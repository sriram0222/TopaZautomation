"""
app.py
------
IT Asset Submission Acknowledgement - single-file desktop app.

Run with:  python app.py

WHY ONE FILE: everything (AD lookup, document filling, signature capture,
cryptographic signing, bulk batch processing, email drafts) lives in this
one script on purpose, so there's exactly one thing to copy, edit, or hand
to another tool - no separate modules, no config.json, no .env file to
keep in sync. All settings are the CONFIG dict right below this docstring
- edit values there directly, save, done.

HOW THIS APP SIGNS DOCUMENTS - NO ADOBE ACROBAT NEEDED, ANYWHERE
This app does the entire job Adobe Acrobat + a Topaz pad would otherwise
do, by itself, end to end:
  1. It captures each signature itself - from a real Topaz SigPlus pad if
     one is connected and enabled (see CONFIG["use_topaz_pad"]), or from
     an on-screen mouse/touch canvas otherwise.
  2. It fills that signature (as a picture) plus every other field into
     the Business Unit's Word template, and exports that to PDF using a
     locally installed Microsoft Word - so the PDF looks exactly like the
     final signed form should, visually.
  3. It then applies TWO REAL, cryptographically signed PDF signature
     fields on top - using pyHanko and a signing certificate this app
     generates for itself the first time it runs (see
     ensure_signing_identity()). This is what makes the file genuinely
     "signed" in the way Acrobat/Reader detect and validate when the file
     is reopened later (Signature Panel, the "Signed and all signatures
     are valid" banner) - not just a picture dropped onto a page, which
     can't be verified or audited that way.

TRADE-OFF, ACCEPTED: the certificate this app signs with is FREE and
SELF-SIGNED (nobody outside this organization vouches for it), so a viewer
will correctly show the document as signed and tamper-evident, but will
also show a caution that the certificate itself isn't in a trust store.
Your IT team can remove that caution later, for everyone, by installing
the certificate this app generates (signing_identity/signing_cert.pem) as
a Trusted Certificate in Acrobat - see the README. Nothing about how the
app works needs to change for that to happen; it only affects how trusted
the existing signatures look.

HOW TO USE IT

1. ONE PERSON AT A TIME (the default): pick Business Unit -> Employee ID
   (AD lookup, manual fallback) -> Contact Number -> Submission Type ->
   conditional asset fields -> Employee Signature -> Asset Receiver
   Signature -> Review (choose a save location if you want one other than
   the default) -> Generate. Generate produces a fully filled AND SIGNED
   PDF, ready to send to an auditor or manager as-is.

2. BULK BATCH (100 new hires, 150 exits, etc.): on the Business Unit
   screen, click "Import Batch from Excel...". You'll see a preview list
   of everyone loaded before anything starts. Once you continue, the app
   auto-advances through everyone - auto-filling Employee ID and Type,
   and auto-running the AD lookup UNLESS the Excel row already has a Name
   (and ideally Manager) - if it does, that row's data is used directly
   and AD Lookup is skipped entirely for that person. This matches how
   exits are typically already fully documented (name, manager, exit
   date all known ahead of time) while new hires/break-fixes usually
   aren't, and still need the live AD lookup.
   The app only stops for you at the parts that genuinely need a human:
   entering serial numbers, capturing both signatures, and confirming
   Generate.
   If a lookup fails or a Type isn't recognized for a particular person,
   THAT one person drops into full manual mode (fix it, proceed as
   normal) - the batch auto-resumes for everyone after them once you hit
   Generate.

ONE-TIME PER-BU SETUP (do this once per Business Unit, before adding it
here): take that BU's base Word (.docx) form and click "Add New BU
Template..." in this app, then pick that file. The app automatically
inserts every fillable field/checkbox/signature control it needs directly
into a copy of that document - you do NOT need to prepare anything in
Word or Acrobat first. If the document's layout doesn't closely match the
expected standard form (the checkbox list, the underline blanks, the
assets table), the app tells you exactly what it couldn't find so you can
fix that spot in the Word doc and re-upload.

ARCHITECTURE NOTE ON THE AD LOOKUP: since your employee lookup portal
relies on your logged-in browser session (SSO) rather than a plain API,
this app drives a real embedded Microsoft Edge window to reuse that
session - opened once, kept open for as long as the app runs. That
embedded browser needs its own event loop, which can't share a process
with the main window's event loop, so it actually runs as a SEPARATE
process - this same script, relaunched with a hidden "--webview-helper"
flag (see the very bottom of this file). You never see this happen; it's
just how "one file, one process model" and "a real embedded browser" both
work at the same time.

WHAT YOU NEED ON THE MACHINE RUNNING THIS
- Python 3.10+
- Microsoft Word (used to export the filled document to PDF - this is the
  only "Office" dependency; Adobe Acrobat/Reader is NOT needed to CREATE
  a signed PDF, only optionally to LOOK at one afterward, same as any
  other PDF viewer)
- A Topaz SigPlus pad is OPTIONAL - only needed if CONFIG["use_topaz_pad"]
  is turned on; otherwise every signature is captured on-screen
- pip install -r requirements.txt
"""

import base64
import io
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
import zipfile
from datetime import date, datetime, timedelta
from html.parser import HTMLParser

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import copy

try:
    import requests
except ImportError:
    requests = None

try:
    from requests_negotiate_sspi import HttpNegotiateAuth
    HAS_SSPI = True
except ImportError:
    HAS_SSPI = False

try:
    import truststore
    HAS_TRUSTSTORE = True
except ImportError:
    HAS_TRUSTSTORE = False

try:
    from reportlab.pdfgen import canvas as _rl_canvas
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False


def _enable_os_trust_store():
    """Make ssl/requests trust the Windows OS certificate store (Trusted Root CAs)
    instead of only the bundled certifi list. Internal corporate root CAs (e.g. a
    company's internal SSRS/reporting servers) are typically installed into the
    OS store by IT policy and trusted by the browser, but NOT present in certifi's
    public CA bundle, causing 'self-signed certificate in certificate chain' SSL
    errors from Python even though the same URL opens fine in Chrome/Edge.
    Safe no-op if the truststore package isn't installed."""
    if not HAS_TRUSTSTORE:
        return False
    try:
        truststore.inject_into_ssl()
        return True
    except Exception:
        return False

from PIL import Image, ImageDraw, ImageFont, ImageTk
import openpyxl


# ============================================================================
# CONFIGURATION - edit these values directly, no separate config file needed
# ============================================================================
CONFIG = {
    "bu_templates_folder": "bu_templates",
    # Where uploaded per-BU Word doc templates are stored (both the
    # original upload and the auto-generated fillable version).

    "ad_lookup_mode": "ssrs_report",
    # "ssrs_report" = for an SSRS/ReportViewer-style "Active Directory
    #             Search" report: an ASP.NET report page with a multi-line
    #             ID textbox and a "View Report" button (not a plain web
    #             form), results shown as a table with column headers like
    #             EmployeeID / Employee Name / UserPrincipalName. No exact
    #             element IDs needed - the app finds the textbox, the
    #             button, and the results table by their ROLE/TEXT instead
    #             of by ID, since SSRS auto-generates long IDs that shift
    #             between deployments and often sit inside an iframe. See
    #             ad_lookup_ssrs below - just set search_page_url.
    # "browser"    = generic version for a plain web form (a real search
    #             box + submit button with stable CSS selectors/IDs).
    # "direct"  = plain HTTP request to a fixed URL pattern, no login
    #             handling (only works if your portal doesn't need SSO).

    "ad_lookup_ssrs": {
        # REQUIRED: the AD lookup report's URL (the "bookmark this URL"
        # link on the report page itself, not the address-bar URL - SSRS
        # report addresses often carry a session-specific piece that
        # doesn't work if reused later. Only ever used locally, never sent
        # anywhere.
        "search_page_url": "REPLACE_ME_with_the_ActiveDirectoryInfo_bookmark_url",
        "view_report_button_text": "View Report",
        # Visible text/label on the button that runs the search - matched
        # case-insensitively, so this rarely needs changing.
        "results_header_keyword": "EmployeeID",
        # Text that appears in the results table's header row once results
        # have loaded - used to detect "results are ready" and to pick the
        # right table if the page has more than one.
        "employee_name_column": "Employee Name",
        # The exact column header text to read the person's name from.
        "results_ready_timeout_seconds": 25,
    },

    "ad_lookup_browser": {
        # Only used if ad_lookup_mode is "browser" (a plain web form, not
        # an SSRS report). REQUIRED in that case: fill these in from your
        # real AD Lookup site. How to get the exact values: open the site
        # in Edge/Chrome, log in normally, right-click the search box ->
        # Inspect -> note its "id". Same for the search button. Do a
        # search, right-click somewhere in the results -> Inspect -> note
        # what wraps them.
        "search_page_url": "http://your-ad-lookup-site.company.com/",
        "search_input_selector": "#REPLACE_ME_search_box_id",
        "search_submit_selector": "#REPLACE_ME_search_button_id",
        "use_enter_key_to_submit": False,
        "results_ready_selector": "#REPLACE_ME_results_container_id",
        "multi_result_row_selector": "",
        "results_ready_timeout_seconds": 20,
    },

    "ad_lookup_url_template": "http://adlookup.company.com/search?empid={emp_id}",
    # Only used if ad_lookup_mode is "direct".
    "ad_lookup_timeout_seconds": 8,

    "submission_type_row1": ["New Hire", "Break Fix", "Mixed Build", "Additional Laptop", "Laptop Return"],
    "submission_type_row2": ["Contractor LWD", "LWD", "Site Transfer", "RCP"],

    # NOTE: earlier versions of this app showed a different subset of the
    # Asset Details fields depending on the submission type (just a "New
    # Device Serial Number" for New Hire, old+new for Break Fix, a full
    # checklist for everything else). The real signed production form
    # always shows every field regardless of submission type, so the
    # wizard now does the same - see _rebuild_asset_details() /
    # _batch_rebuild_asset_details().

    "asset_checklist_items": ["Laptop Bag", "Charger", "Mouse", "Headset", "Network Cable", "Multiport"],

    "bulk_type_aliases": {
        # Maps short/informal 'Type' values from a bulk-import Excel file
        # to the app's real submission types. Add more synonyms any time -
        # matching is case-insensitive. "ITP" is deliberately NOT mapped
        # here yet - it doesn't match any of the 9 real submission types,
        # so those rows get flagged for a manual pick until it's clear
        # what it should map to.
        "New Hire": ["newhire", "new hire", "new join", "joiner", "new"],
        "Break Fix": ["breakfix", "break fix", "repair", "fix"],
        "Mixed Build": ["mixedbuild", "mixed build"],
        "Additional Laptop": ["additionallaptop", "additional laptop", "extra laptop"],
        "Laptop Return": ["laptopreturn", "laptop return"],
        "Contractor LWD": ["contractorlwd", "contractor lwd", "contractor exit"],
        "LWD": ["exit", "lwd", "leaver", "resignation", "resign", "separation"],
        "Site Transfer": ["sitetransfer", "site transfer", "transfer"],
        "RCP": ["rcp"],
    },

    "email_draft_submission_types": ["LWD", "Contractor LWD"],
    # After generating the PDF for one of these submission types, a
    # pre-filled Outlook draft (Emp ID/Name/assets returned, PDF attached)
    # opens for review - it is NEVER sent automatically. Add "Laptop
    # Return", "Site Transfer", "RCP" here too if you want the same.
    "lwd_email_to": "",
    # Optional default "To" address for that draft. Leave blank to fill
    # it in yourself each time.

    "pdf_save_folder": "C:\\AssetForms\\Generated",
    # Where generated (signed) PDFs are saved by default (the "Choose
    # Location..." button on the review screen can override this per
    # form). Each submission type gets its own subfolder underneath this
    # automatically (e.g. ...\\Generated\\Break_Fix\\, ...\\Generated\\LWD\\).

    "tracit_submission_types": ["Break Fix", "LWD", "Contractor LWD"],
    # Submission types that show an "Open TracIT" button next to the
    # Current Device Serial Number field.

    "no_current_asset_submission_types": ["New Hire"],
    # Submission types where the person has no existing/old device, so
    # the "Current Device Serial Number" field is hidden entirely (only
    # "New Device Serial Number" applies). New Hire is the obvious case;
    # add others here if they ever come up.
    "tracit_url_template": "https://tracit.optum.com/ham/view-assets",
    # The TracIT page to open. This is the base page only - the exact
    # query-string format for deep-linking straight to an employee ID or
    # serial number search wasn't available yet, so the button opens this
    # page as-is and the ID/serial are typed in there by hand, same as
    # today. If TracIT DOES support a URL format like
    # "https://tracit.optum.com/ham/view-assets?empId={employee_id}&serial={serial_number}",
    # paste that real (working) URL here instead - the {employee_id} and
    # {serial_number} placeholders will be filled in automatically from
    # the form.

    "signing_identity_folder": "signing_identity",
    # Where this app's own self-signed signing certificate + private key
    # are stored, once generated on first use (see ensure_signing_identity
    # below). Keep this folder private - anyone with the private key file
    # could produce signatures that claim to be from this app.

    "use_topaz_pad": True,
    # True = capture both signatures from a real, connected Topaz SigPlus
    # pad. False (default) = always use the on-screen mouse/touch canvas -
    # this is what a machine with no signature pad attached (e.g. for
    # testing) should use. If a Topaz pad is enabled but not actually
    # available at the moment, the app automatically falls back to the
    # on-screen canvas for that signature rather than getting stuck.
    "topaz_progid_options": ["TOPAZSIGPLUS.SigPlusCtrl.1", "SigPlus.SigPlusCtrl.1"],
    # COM ProgIDs to try, in order, when connecting to the Topaz SDK.

    "asset_receiver_display_name": "Asset Receiver",
    # Name recorded against the second cryptographic signature (the
    # employee's own name is used automatically for the first).

    "signing_reason": "IT Asset Acknowledgement",
    # Recorded as the cryptographic signature's "Reason" (shown in Adobe
    # Acrobat's Signature Properties panel for both signature fields).
    # This does NOT make the certificate itself "Trusted" in Acrobat -
    # that still requires your IT team to install it as a Trusted
    # Certificate (see the README) - it only adds readable context to an
    # otherwise-unlabeled signature.

    "email_templates": {
        # Per-submission-type label used in the pre-filled Outlook draft's
        # subject/body (see build_email_body). Falls back to the
        # submission type's own name if not listed here. Only types also
        # listed in "email_draft_submission_types" above actually get a
        # draft opened.
        "New Hire": "Asset Issued",
        "Break Fix": "Replacement Asset Issued",
        "LWD": "Asset Returned",
        "Contractor LWD": "Asset Returned",
        "Site Transfer": "Asset Transfer",
    },

    "ssrs_asset_report": {
        # The SSRS report this app can download (via the current Windows
        # login session - see requests_negotiate_sspi above, no stored
        # password) to auto-populate New Hire asset details from Employee
        # ID, instead of typing them in by hand every time.
        "url": (
            "https://orbit-ssrs-prod-int.optum.com/ReportServer?/Optum%20ET/"
            "EUTS%20Dev%20Ops/EUTSDReporting/TransferredFromPowerBIReportServer/"
            "CheckInTool-ShippingCenterReports/DDT%20Open%20Requests%20by%20"
            "Status%20Raw%20Data&rs:Format=EXCELOPENXML"
        ),
        # Only re-downloaded once per calendar day automatically (see
        # Enhancement 9) - "Refresh Data" in the app always forces a fresh
        # download regardless of this cache.
        "cache_max_age_hours": 24,
        # Submission types where SSRS is the PRIMARY lookup source (checked
        # automatically the moment an Employee ID is entered - see
        # _ssrs_autocheck_now/_batch_try_ssrs_autofill_identity) - AD is
        # only used as a fallback for these, e.g. when SSRS has no manager
        # name (SSRS's column map doesn't include one). LWD/Contractor
        # LWD/Site Transfer are intentionally NOT in this list - for those,
        # AD + the imported Excel data remain the (only) lookup sources.
        "new_hire_submission_types": ["New Hire", "Break Fix", "Mixed Build", "Additional Laptop"],
        # SSRS column name -> app field. Edit the right-hand column names
        # here if the real report's headers differ slightly (e.g. a
        # trailing space, or a renamed column) - nothing else needs to
        # change.
        "column_map": {
            "employee_id": "Empl ID",
            "employee_name": "Recipient Name",
            "email": "Email Address",
            "serial_number": "New SN",
            "hostname": "Hostname New",
            "device_model": "Described Model New",
            "ticket_number": "Prob Ticket",
            "monitor": "Monitor Type",
            "dock": "Peripherals Dock",
            "keyboard_mouse": "Peripherals Keyb Mouse",
            "power_adapter": "Peripherals Power Chords",
            "battery": "Peripherals Battery",
        },
    },

    "servicenow": {"enabled": False},
    # Placeholder for a future ServiceNow integration - not implemented
    # yet, kept here so the shape of that config already exists.
}


SUBMISSION_TYPE_TAGS = {
    "New Hire": "NewHire", "Break Fix": "BreakFix", "Mixed Build": "MixedBuild",
    "Additional Laptop": "AdditionalLaptop", "Laptop Return": "LaptopReturn",
    "Contractor LWD": "ContractorLWD", "LWD": "LWD", "Site Transfer": "SiteTransfer",
    "RCP": "RCP",
}
ASSET_CHECKLIST_TAGS = {
    "Laptop Bag": "LaptopBag", "Charger": "Charger", "Mouse": "Mouse",
    "Headset": "Headset", "Network Cable": "NetworkCable", "Multiport": "Multiport",
}


def _app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ============================================================================
# DIAGNOSTIC LOGGING - writes a plain-text log file next to app.py so a
# problem in the field (AD lookup, Topaz pad, PDF generation, etc.) can be
# diagnosed from a screenshot/copy of that file, instead of needing a
# screen recording every time. Safe to leave on permanently - it rotates
# once the file gets large rather than growing forever.
# ============================================================================
import logging
import logging.handlers

logger = logging.getLogger("it_asset_app")


def _setup_logging():
    log_dir = os.path.join(_app_dir(), "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, "application.log")

    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    if logger.handlers:
        return log_path  # already set up (e.g. re-entrant call)

    fmt = logging.Formatter(
        "%(asctime)s %(levelname)-7s [%(threadName)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.handlers.RotatingFileHandler(
        log_path, maxBytes=2_000_000, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    file_handler.setLevel(logging.DEBUG)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(fmt)
    console_handler.setLevel(logging.INFO)
    logger.addHandler(console_handler)

    logger.info("=" * 70)
    logger.info("App starting - log file: %s", log_path)
    logger.info("Python: %s | frozen: %s", sys.version.split()[0], getattr(sys, "frozen", False))

    def _log_uncaught(exc_type, exc_value, exc_tb):
        logger.critical("UNCAUGHT EXCEPTION", exc_info=(exc_type, exc_value, exc_tb))
        sys.__excepthook__(exc_type, exc_value, exc_tb)

    sys.excepthook = _log_uncaught
    return log_path


# ============================================================================
# AD LOOKUP - direct HTTP mode (used when ad_lookup_mode = "direct")
# ============================================================================

# Maps our internal field name -> list of possible label strings the AD page
# might use for that field. Matching is case-insensitive and ignores extra
# whitespace/colons. Add more synonyms here if your portal uses different
# wording.
FIELD_LABEL_MAP = {
    "emp_name": ["employee name", "name", "full name", "emp name"],
    "manager_name": ["manager", "manager name", "reporting to", "reports to", "reporting manager"],
    "contact_number": ["contact number", "phone", "mobile", "contact no", "phone number"],
    "department": ["department", "dept"],
    "email": ["email", "email id", "e-mail"],
    "designation": ["designation", "title", "job title"],
}


class _TextExtractor(HTMLParser):
    """Very small HTML->text converter (no external deps required).
    Keeps line breaks so label/value pairs that sit in separate table
    cells or <div>/<p> blocks stay on separate lines, which makes the
    regex-based field extraction below much more reliable."""

    BLOCK_TAGS = {"tr", "div", "p", "br", "li", "td", "th"}

    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag in self.BLOCK_TAGS:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in self.BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data):
        if data and data.strip():
            self.parts.append(data.strip())

    def get_text(self):
        return "\n".join(p for p in self.parts if p)


def _html_to_lines(html):
    extractor = _TextExtractor()
    extractor.feed(html)
    text = extractor.get_text()
    # collapse multiple blank lines
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    return lines


def _extract_fields(lines):
    """
    Two extraction strategies, tried in order for every field:
      1. Same-line pattern: "Label: Value" or "Label - Value"
      2. Next-line pattern: label is its own line, value is the line after
         (common when a lookup page renders label/value in separate table
         cells that end up on consecutive lines after HTML stripping).
    """
    result = {key: "" for key in FIELD_LABEL_MAP}
    joined = "\n".join(lines)

    for field, labels in FIELD_LABEL_MAP.items():
        for label in labels:
            # Strategy 1: same line, "Label : Value"
            pattern = re.compile(
                re.escape(label) + r"\s*[:\-]\s*(.+)", re.IGNORECASE
            )
            match = pattern.search(joined)
            if match:
                value = match.group(1).strip()
                # cut off if it accidentally grabbed the next label too
                value = value.split("\n")[0].strip()
                if value:
                    result[field] = value
                    break

            # Strategy 2: label alone on a line, value on the next line
            for i, line in enumerate(lines):
                if line.strip().lower().rstrip(":").strip() == label.lower():
                    if i + 1 < len(lines):
                        candidate = lines[i + 1].strip()
                        if candidate:
                            result[field] = candidate
                            break
            if result[field]:
                break

    return result


def fetch_employee_details(emp_id, config):
    """
    Looks up an employee by ID against the configured AD Lookup URL.

    Returns: dict with keys emp_name, manager_name, contact_number,
             department, email, designation (each "" if not found)
    Raises: RuntimeError with a human-readable message on failure, so the
            calling GUI code can show it directly to the user.
    """
    if requests is None:
        raise RuntimeError(
            "The 'requests' library is not installed. Run:\n"
            "    pip install requests\n"
            "(and optionally 'pip install requests-negotiate-sspi' if your "
            "AD Lookup page needs Windows network authentication)."
        )

    emp_id = (emp_id or "").strip()
    if not emp_id:
        raise RuntimeError("Please enter an Employee ID before looking it up.")

    url_template = config.get("ad_lookup_url_template", "")
    if "{emp_id}" not in url_template:
        raise RuntimeError(
            "config.json -> ad_lookup_url_template must contain the "
            "placeholder {emp_id}, e.g. "
            "'http://adlookup.company.com/search?id={emp_id}'"
        )

    url = url_template.format(emp_id=emp_id)
    timeout = config.get("ad_lookup_timeout_seconds", 8)

    auth = HttpNegotiateAuth() if HAS_SSPI else None

    try:
        resp = requests.get(url, timeout=timeout, auth=auth)
        resp.raise_for_status()
    except requests.exceptions.ConnectionError:
        raise RuntimeError(
            f"Could not reach the AD Lookup page at:\n{url}\n\n"
            "Check that you're connected to the corporate network/VPN."
        )
    except requests.exceptions.Timeout:
        raise RuntimeError("The AD Lookup page took too long to respond.")
    except Exception as exc:
        raise RuntimeError(f"AD Lookup failed: {exc}")

    lines = _html_to_lines(resp.text)
    fields = _extract_fields(lines)

    if not any(fields.values()):
        raise RuntimeError(
            "Connected to the AD Lookup page, but couldn't find any known "
            "fields in the response. Open ad_lookup.py and update "
            "FIELD_LABEL_MAP with the exact label text your portal uses "
            f"for Employee ID '{emp_id}'."
        )

    return fields


def _uses_webview_lookup(config):
    """True for any ad_lookup_mode that drives the embedded-browser helper
    process (BrowserLookupSession) rather than a plain HTTP GET. Both
    "browser" (a generic web form) and "ssrs_report" (an SSRS ReportViewer
    report) need the real embedded browser window for SSO/session reuse -
    only "direct" mode skips it."""
    return config.get("ad_lookup_mode", "browser") in ("browser", "ssrs_report")


# ============================================================================
# AD LOOKUP - persistent embedded-browser session (client side, runs in
# the main GUI process, talks to the helper process below over stdin/stdout)
# ============================================================================

class BrowserLookupSession:
    SEARCH_TIMEOUT_SECONDS = 60

    def __init__(self):
        self._proc = None
        self._lock = threading.Lock()
        self._pending = {}  # request_id -> {"event": Event, "result": dict}
        self._fatal_error = None

    def _build_command(self):
        # Single-file app: the helper isn't a separate script/exe anymore -
        # this just re-launches THIS SAME program with a special flag, and
        # the dispatcher at the bottom of the file routes to the webview
        # helper instead of the GUI when it sees that flag.
        if getattr(sys, "frozen", False):
            return [sys.executable, "--webview-helper"]
        return [sys.executable, os.path.abspath(__file__), "--webview-helper"]

    def start(self):
        """Launches the helper process and its background window. Safe to
        call more than once - a no-op if already running."""
        with self._lock:
            if self._proc is not None and self._proc.poll() is None:
                # NOTE: this is a no-op reuse of an already-running helper
                # subprocess. If app.py was just replaced on disk (a new
                # build/fix copied over the old one) but this app was never
                # fully closed and relaunched, the helper here is still the
                # OLD process running the OLD code - a rebuilt/fixed app.py
                # will NOT take effect for AD lookups until the whole app
                # is closed (check Task Manager for a lingering python.exe/
                # IT_Asset_Form.exe helper) and reopened fresh.
                logger.debug(
                    "AD helper: reusing already-running helper subprocess (pid=%s) - "
                    "if app.py was just updated, this process is still on the OLD code "
                    "until the app is fully closed and reopened.",
                    self._proc.pid,
                )
                return
            self._fatal_error = None
            command = self._build_command()
            try:
                self._proc = subprocess.Popen(
                    command,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    bufsize=1,
                )
                logger.info("AD helper: launched new subprocess (pid=%s): %r", self._proc.pid, command)
            except FileNotFoundError:
                logger.exception("AD helper: failed to launch subprocess: %r", command)
                self._fatal_error = (
                    f"Could not launch the AD lookup helper process ({command[0]}). "
                    "This usually means Python itself couldn't be found - "
                    "if you're running the packaged .exe, try reinstalling it."
                )
                self._proc = None
                return

            threading.Thread(target=self._read_loop, daemon=True).start()

    def _read_loop(self):
        proc = self._proc
        for line in proc.stdout:
            line = line.strip()
            if not line:
                continue
            try:
                data = json.loads(line)
            except Exception:
                continue

            if data.get("cmd") == "fatal_error":
                self._fatal_error = data.get("error", "Unknown fatal error in lookup helper.")
                # Resolve any request(s) already waiting - e.g. the very
                # first search() call, which races start() and can write
                # its request before this fatal_error arrives. Without
                # this, that call's Event never gets set and it sits for
                # the full SEARCH_TIMEOUT_SECONDS before timing out with a
                # generic message instead of surfacing this one instantly.
                for entry in self._pending.values():
                    entry["result"] = {"error": self._fatal_error}
                    entry["event"].set()
                continue

            request_id = data.get("request_id")
            entry = self._pending.get(request_id)
            if entry is not None:
                entry["result"] = data
                entry["event"].set()
        # stdout closed -> helper process ended (crashed, or window closed
        # by the user). Mark it so the next search() call restarts it.
        with self._lock:
            if self._proc is proc:
                self._proc = None

    def is_running(self):
        return self._proc is not None and self._proc.poll() is None

    def search(self, emp_id):
        """Blocks until candidates are returned or raises RuntimeError.
        Returns a list of field dicts (length 1 = single match, >1 = the
        caller should ask the user to pick one, 0 never returned - a
        no-match search should still raise or return an empty list from
        the helper's own parsing, surfaced to the caller as an empty list)."""
        if not self.is_running():
            self.start()
        if self._fatal_error:
            raise RuntimeError(self._fatal_error)
        if self._proc is None:
            raise RuntimeError("Lookup helper failed to start.")

        request_id = uuid.uuid4().hex
        event = threading.Event()
        self._pending[request_id] = {"event": event, "result": None}

        try:
            self._proc.stdin.write(json.dumps({"request_id": request_id, "emp_id": emp_id}) + "\n")
            self._proc.stdin.flush()
        except Exception as exc:
            self._pending.pop(request_id, None)
            raise RuntimeError(
                f"Lookup helper is not responding ({exc}). It may have been "
                "closed - click Lookup again to restart it."
            )

        if not event.wait(timeout=self.SEARCH_TIMEOUT_SECONDS):
            self._pending.pop(request_id, None)
            raise RuntimeError(
                "Timed out waiting for the AD Lookup search result. If a "
                "login page appeared, please complete it and try again."
            )

        result = self._pending.pop(request_id)["result"]
        if "error" in result:
            raise RuntimeError(result["error"])
        return result.get("candidates", [])

    def shutdown(self):
        """Call when the main app is closing, to cleanly close the
        embedded browser window and end the helper process."""
        with self._lock:
            if self._proc is None:
                return
            try:
                self._proc.stdin.write(json.dumps({"cmd": "shutdown"}) + "\n")
                self._proc.stdin.flush()
            except Exception:
                pass
            try:
                self._proc.wait(timeout=5)
            except Exception:
                try:
                    self._proc.terminate()
                except Exception:
                    pass
            self._proc = None


# ============================================================================
# AD LOOKUP - embedded-browser helper (runs in a SEPARATE process - this
# same file, relaunched with --webview-helper; see the bottom of this file)
# ============================================================================

STDOUT_LOCK = threading.Lock()


def _send(obj):
    with STDOUT_LOCK:
        sys.stdout.write(json.dumps(obj) + "\n")
        sys.stdout.flush()


# ----------------------------------------------------------------------
# SSRS/ReportViewer-style AD lookup (ad_lookup_mode = "ssrs_report").
# Finds the parameter textbox, the "View Report" button, and the results
# table by ROLE/TEXT rather than by ID - SSRS auto-generates long element
# IDs that shift between deployments, and the report often renders inside
# an iframe, so a fixed CSS selector is the wrong tool here.
# ----------------------------------------------------------------------

_JS_COLLECT_DOCS = """
function __collectDocs() {
    // Walks the ENTIRE iframe tree, not just the top document's direct
    // children. The real SSRS/ReportViewer page nests iframes at least
    // two deep - a report-frame iframe hangs off the main page (that's
    // where the parameter textbox and "View Report" button live), and
    // the actual results table renders inside a SEPARATE iframe nested
    // inside THAT one. A one-level-only walk finds the textbox/button
    // fine (both are one level down) but never sees the results table
    // at all - which is exactly what field evidence showed: a screen
    // recording with the real row plainly visible on screen, while the
    // app timed out never finding a single matching table. A depth cap
    // guards against a pathological/circular frame tree.
    var docs = [document];
    var seen = new Set();
    function walk(doc, depth) {
        if (!doc || depth > 6 || seen.has(doc)) return;
        seen.add(doc);
        try {
            var iframes = doc.querySelectorAll('iframe, frame');
            for (var i = 0; i < iframes.length; i++) {
                try {
                    var d = iframes[i].contentDocument;
                    if (d && !seen.has(d)) {
                        docs.push(d);
                        walk(d, depth + 1);
                    }
                } catch (e) { /* cross-origin iframe, skip */ }
            }
        } catch (e) {}
    }
    walk(document, 0);
    return docs;
}
"""

# ----------------------------------------------------------------------
# ONE shared table scanner used by BOTH the "are the results ready yet?"
# poll and the "now pull the rows out" extraction below.
#
# It is deliberately a single function so the two steps can never
# disagree about WHICH table on the page is the results table - which is
# exactly the bug this replaced. The old code had two separate loops:
# the readiness check accepted only a matching table with >1 rows, while
# the extractor returned from the FIRST matching table it saw whether or
# not it had any data rows. SSRS/ReportViewer renders the same report as
# several nested tables (an outer layout table wrapping the real one, and
# frequently a separate "fixed header" table holding just the header row,
# with the data rows in a table of their own). So the check would happily
# find the real data table and report "ready", then the extractor would
# hit the header-only/layout table first, find zero data rows under it,
# and give up - producing "The report loaded, but returned no rows" for
# an employee who is plainly right there on screen.
#
# The scanner therefore looks at EVERY table, and:
#   1. prefers a matching table that carries its own data rows (picking
#      the one with the most, so an outer wrapper never wins over the
#      real inner table), and
#   2. falls back to pairing a header-only matching table with the next
#      table that actually holds rows, for the split-header rendering.
# It also returns what it saw in every table, so a failure gets logged
# with the real column headers instead of a bare "no rows".
# ----------------------------------------------------------------------
_JS_SCAN_TABLES = """
function __rowTexts(row) {
    var out = [];
    if (!row || !row.cells) return out;
    for (var i = 0; i < row.cells.length; i++) {
        out.push((row.cells[i].innerText || row.cells[i].textContent || '').trim());
    }
    return out;
}

function __mapRow(headers, cellTexts) {
    // The real report's header row has "Department" as a column twice
    // (confirmed against an actual exported copy of the page) - keep
    // the FIRST occurrence's value rather than letting the second
    // silently overwrite it, so a duplicate column name can't quietly
    // drop real data.
    var obj = {};
    var n = Math.min(headers.length, cellTexts.length);
    for (var i = 0; i < n; i++) {
        if (headers[i] && !(headers[i] in obj)) obj[headers[i]] = cellTexts[i];
    }
    return obj;
}

// A real results header row looks like "EmployeeID | Employee Name |
// Manager Name | ..." - several SHORT column labels. It is NOT a single
// cell containing a wall of concatenated text. That distinction matters
// because the ReportViewer's own toolbar/chrome (the "Search AD for
// MSID, EmployeeID, or Email Address" label, the zoom/export dropdowns,
// the page-count widget, "Loading...") also renders as <table> markup on
// the page, and that descriptive label text happens to CONTAIN the
// header keyword ("EmployeeID") as a substring - so without this guard
// the toolbar itself gets mistaken for the results table, especially in
// the first instant after clicking "View Report" while the real report
// is still loading. That produced a report that looked "ready" after
// well under a second and then extracted zero real rows - for an
// employee who was genuinely in AD - because it was reading the
// toolbar, not the results.
function __looksLikeHeaderRow(headers) {
    if (headers.length < 2) return false;
    for (var i = 0; i < headers.length; i++) {
        if (!headers[i] || headers[i].length > 120) return false;
    }
    return true;
}

// SSRS-exported HTML tables commonly lead with a "sizer" row - a row of
// mostly-empty cells whose only job is to fix column widths (colgroup
// via <tr><td style="width:Npx">) - BEFORE the real header row. The
// scanner used to only ever look at t.rows[0], so on any table shaped
// that way it read the blank sizer row, __looksLikeHeaderRow correctly
// rejected it (blank cells), and the table was never picked - even
// though the real header (and the data under it) was sitting right
// there one row down. This looks for the header among the first few
// rows of a table instead of assuming it's always row 0.
function __findHeaderRowIndex(t, lower) {
    var maxCheck = Math.min(t.rows ? t.rows.length : 0, 4);
    for (var idx = 0; idx < maxCheck; idx++) {
        var headers = __rowTexts(t.rows[idx]);
        if (!__looksLikeHeaderRow(headers)) continue;
        for (var k = 0; k < headers.length; k++) {
            if (headers[k].toLowerCase().indexOf(lower) !== -1) return idx;
        }
    }
    return -1;
}

function __scanTables(headerKeyword) {
    var lower = String(headerKeyword).toLowerCase();
    var docs = __collectDocs();
    var all = [];
    for (var i = 0; i < docs.length; i++) {
        var tables = docs[i].querySelectorAll('table');
        for (var j = 0; j < tables.length; j++) {
            var t = tables[j];
            var totalRows = t.rows ? t.rows.length : 0;
            var headerIdx = __findHeaderRowIndex(t, lower);
            var matched = headerIdx !== -1;
            var headers = matched ? __rowTexts(t.rows[headerIdx]) : __rowTexts(t.rows[0]);
            // "rowCount" here means rows of potential DATA available
            // under the header we actually found (not the table's raw
            // row count) - so a table with a sizer row + header + 5 data
            // rows correctly scores as 5, not 7, and a header-only table
            // (nothing but the header, real data rendered in a separate
            // table entirely) still scores as 0 rather than 1.
            var dataRowCount = matched ? (totalRows - headerIdx - 1) : totalRows;
            all.push({ table: t, headers: headers, matched: matched,
                       headerIdx: matched ? headerIdx : 0, rowCount: dataRowCount });
        }
    }

    var summary = [];
    for (var s = 0; s < all.length; s++) {
        summary.push({ headers: all[s].headers, rowCount: all[s].rowCount,
                       matched: all[s].matched });
    }

    // 1. Best case: a matching table that holds its own data rows. Pick
    //    the one with the MOST rows so an outer wrapper table (which can
    //    also "match", because the real header text is nested inside it)
    //    never beats the real inner results table.
    var best = null;
    for (var a = 0; a < all.length; a++) {
        if (!all[a].matched || all[a].rowCount < 1) continue;
        if (best === null || all[a].rowCount > best.rowCount) best = all[a];
    }
    if (best !== null) {
        var rows = [];
        for (var r = best.headerIdx + 1; r < best.table.rows.length; r++) {
            var texts = __rowTexts(best.table.rows[r]);
            if (!texts.length) continue;
            var isBlank = true;
            for (var b = 0; b < texts.length; b++) { if (texts[b]) { isBlank = false; break; } }
            if (isBlank) continue;
            rows.push(__mapRow(best.headers, texts));
        }
        if (rows.length) {
            return { rows: rows, tables: summary, strategy: 'single-table' };
        }
    }

    // 2. Split rendering: the header row lives in its own table (possibly
    //    with a sizer row above it) and the data rows are in a later
    //    table. Pair them up.
    for (var h = 0; h < all.length; h++) {
        if (!all[h].matched || all[h].rowCount !== 0) continue;
        var headers2 = all[h].headers;
        if (!headers2.length) continue;
        for (var d = h + 1; d < all.length; d++) {
            var cand = all[d];
            if (!cand.table.rows || !cand.table.rows.length) continue;
            var firstTexts = __rowTexts(cand.table.rows[0]);
            if (!firstTexts.length) continue;
            // Prefer an exact column-count match; that is almost always
            // the real data table for this header.
            if (firstTexts.length !== headers2.length) continue;
            var rows2 = [];
            for (var r2 = 0; r2 < cand.table.rows.length; r2++) {
                var texts2 = __rowTexts(cand.table.rows[r2]);
                if (!texts2.length) continue;
                var blank2 = true;
                for (var b2 = 0; b2 < texts2.length; b2++) { if (texts2[b2]) { blank2 = false; break; } }
                if (blank2) continue;
                // Skip a repeat of the header row itself.
                if (texts2.join('\\u0001') === headers2.join('\\u0001')) continue;
                rows2.push(__mapRow(headers2, texts2));
            }
            if (rows2.length) {
                return { rows: rows2, tables: summary, strategy: 'split-header' };
            }
        }
    }

    return { rows: [], tables: summary, strategy: 'none' };
}
"""


def _map_ssrs_row_to_fields(row, name_column):
    """Maps one raw {column_header: cell_text} row from the SSRS AD Search
    report into the app's internal field-name dict.

    NOTE: this used to assume Manager wasn't a column on this report at
    all, and left manager_name permanently blank as a result - but the
    real report does have one ("Manager Name", alongside Manager
    EmployeeID/MSID/Email), so that was silently dropping it on every
    successful lookup. Fixed below."""
    lower_row = {(k or "").strip().lower(): (v or "").strip() for k, v in row.items()}

    def get_ci(*keys):
        for key in keys:
            value = lower_row.get(key.lower())
            if value:
                return value
        return ""

    fields = {key: "" for key in FIELD_LABEL_MAP}
    fields["emp_name"] = get_ci(name_column, "employee name", "name")
    fields["manager_name"] = get_ci("manager name", "manager", "reporting manager", "reports to")
    fields["email"] = get_ci(
        "corporate email box", "userprincipalname", "uht-identitymanagement-mail", "email"
    )
    fields["designation"] = get_ci("employee type")
    fields["department"] = get_ci("department")
    # Extra columns from this specific report - not part of FIELD_LABEL_MAP
    # and not shown on the wizard screen today, but harmless to carry
    # along in case a future screen wants them.
    fields["msid"] = get_ci("msid")
    fields["employee_id_confirmed"] = get_ci("employeeid")
    fields["manager_email"] = get_ci("manager email")
    fields["manager_msid"] = get_ci("manager msid")
    fields["manager_employee_id"] = get_ci("manager employeeid")
    fields["account_enabled"] = get_ci("enabled")
    return fields


def _perform_ssrs_search(window, emp_id, button_text, header_keyword, name_column, ready_timeout):
    """Runs one search against an SSRS 'Active Directory Search' report:
    fills the parameter textbox, clicks the button named by `button_text`,
    waits for a results table whose header row contains `header_keyword`,
    then extracts it."""
    logger.info("AD lookup: starting search for emp_id=%r (header_keyword=%r, name_column=%r)",
                emp_id, header_keyword, name_column)

    fill_js = _JS_COLLECT_DOCS + f"""
    (function(empId) {{
        var docs = __collectDocs();
        for (var i = 0; i < docs.length; i++) {{
            var ta = docs[i].querySelector('textarea');
            if (ta) {{
                ta.focus();
                ta.value = empId;
                ta.dispatchEvent(new Event('input', {{ bubbles: true }}));
                ta.dispatchEvent(new Event('change', {{ bubbles: true }}));
                ta.blur();
                return "OK";
            }}
        }}
        return "NO_INPUT";
    }})({json.dumps(emp_id)});
    """
    if window.evaluate_js(fill_js) != "OK":
        logger.error("AD lookup: could not find the parameter textbox for emp_id=%r", emp_id)
        raise RuntimeError(
            "Could not find the employee-ID entry box on the AD lookup "
            "report page (looked for the report's parameter textbox). "
            "Check that ad_lookup_ssrs.search_page_url in CONFIG opens "
            "straight to the 'Active Directory Search' report."
        )
    logger.debug("AD lookup: filled search textbox OK")

    time.sleep(0.4)  # give any onchange/validation JS a moment to run

    click_js = _JS_COLLECT_DOCS + f"""
    (function(buttonText) {{
        var docs = __collectDocs();
        var lower = buttonText.toLowerCase();
        for (var i = 0; i < docs.length; i++) {{
            var candidates = docs[i].querySelectorAll(
                'input[type="submit"], input[type="button"], button, a'
            );
            for (var j = 0; j < candidates.length; j++) {{
                var el = candidates[j];
                var label = (
                    el.value || el.innerText || el.textContent || el.title || ''
                ).toLowerCase();
                if (label.indexOf(lower) !== -1) {{
                    el.click();
                    return "OK";
                }}
            }}
        }}
        return "NO_BUTTON";
    }})({json.dumps(button_text)});
    """
    if window.evaluate_js(click_js) != "OK":
        logger.error("AD lookup: could not find button labeled %r", button_text)
        raise RuntimeError(
            f"Could not find a button/link labeled '{button_text}' on the "
            "AD lookup report page. If your report uses different wording, "
            "update ad_lookup_ssrs.view_report_button_text in CONFIG."
        )
    logger.debug("AD lookup: clicked '%s' button OK", button_text)

    # Require a matching header row AND at least one data row underneath it
    # before calling the report "ready" - not just a matching header on its
    # own. SSRS's report viewer can render the header shell an instant
    # before the data rows actually populate underneath it; checking for
    # the header alone let this poll return "ready" a beat too early, so
    # the extraction step right after it would find 0 data rows even
    # though the employee genuinely exists (visible a moment later if you
    # look at the report by hand). Waiting for a real data row here closes
    # that race without changing behavior for a truly-not-found ID, which
    # still correctly times out below.
    # "Ready" is now defined as "the scanner can actually pull rows out",
    # using the very same scanner the extraction step uses - so the two
    # can no longer pick different tables and disagree.
    check_js = _JS_COLLECT_DOCS + _JS_SCAN_TABLES + f"""
    (function(kw) {{ return __scanTables(kw).rows.length > 0; }})({json.dumps(header_keyword)});
    """
    diag_js = _JS_COLLECT_DOCS + _JS_SCAN_TABLES + f"""
    (function(kw) {{ return __scanTables(kw); }})({json.dumps(header_keyword)});
    """
    wait_start = time.time()
    deadline = wait_start + ready_timeout
    found = False
    js_error = None
    while time.time() < deadline:
        try:
            if window.evaluate_js(check_js):
                found = True
                break
        except Exception as exc:
            # A JS exception inside check_js (e.g. a cross-origin frame
            # access that only fails once execution actually touches it)
            # would otherwise look EXACTLY like "nothing ever matched" -
            # a bare timeout with zero clue why. Surface it instead of
            # silently retrying into the same failure for the whole
            # timeout window.
            js_error = exc
            logger.debug("AD lookup: check_js raised (will keep polling): %s", exc)
        time.sleep(0.4)
    if not found:
        # Always capture what the scanner actually saw at the moment of
        # giving up - previously this diagnostic dump only happened on
        # the separate "ready but 0 rows extracted" path, so a hard
        # timeout (which is what a real header-row-detection bug
        # produces) logged nothing more useful than "timed out", making
        # it impossible to tell a genuinely-missing employee apart from
        # a scanner bug without another round trip.
        try:
            timeout_scan = window.evaluate_js(diag_js) or {}
        except Exception as exc:
            timeout_scan = {}
            if js_error is None:
                js_error = exc
        logger.error(
            "AD lookup: TIMED OUT after %.1fs waiting for results table "
            "(header_keyword=%r) for emp_id=%r%s. Tables seen on the page "
            "at timeout: %r",
            time.time() - wait_start, header_keyword, emp_id,
            f" (check_js raised: {js_error})" if js_error is not None else "",
            timeout_scan.get("tables"),
        )
        raise RuntimeError(
            "Timed out waiting for the AD lookup report results to load "
            f"(looked for a table with '{header_keyword}' in its header "
            f"row, with at least one data row under it). Employee ID "
            f"'{emp_id}' may not exist, or the report is slow to render - "
            "try increasing ad_lookup_ssrs.results_ready_timeout_seconds "
            "in CONFIG."
        )
    logger.debug("AD lookup: results table ready after %.1fs", time.time() - wait_start)

    extract_js = _JS_COLLECT_DOCS + _JS_SCAN_TABLES + f"""
    (function(kw) {{ return __scanTables(kw); }})({json.dumps(header_keyword)});
    """
    # Belt-and-suspenders against the same race the check above guards
    # against: if the table got re-rendered in the instant between the
    # check succeeding and this extraction running, retry a few times
    # before giving up, rather than reporting "no rows" on one bad read.
    rows = []
    scan = {}
    for _attempt in range(5):
        scan = window.evaluate_js(extract_js) or {}
        rows = scan.get("rows") or []
        if rows:
            break
        time.sleep(0.4)

    logger.debug(
        "AD lookup: extracted %d row(s) via strategy=%r; tables seen on page: %r",
        len(rows), scan.get("strategy"), scan.get("tables"),
    )
    logger.debug("AD lookup: raw extracted rows: %r", rows)

    candidates = []
    for row in rows:
        fields = _map_ssrs_row_to_fields(row, name_column)
        logger.debug("AD lookup: mapped fields from row: %r", fields)
        if any(fields.values()):
            candidates.append(fields)

    if not candidates:
        logger.error(
            "AD lookup: report loaded but 0 usable candidates for emp_id=%r "
            "(raw rows=%d, strategy=%r). Tables seen on the page were: %r",
            emp_id, len(rows), scan.get("strategy"), scan.get("tables"),
        )
        raise RuntimeError(
            f"The report loaded, but returned no rows for Employee ID "
            f"'{emp_id}'. Double-check the ID is correct and enabled in AD."
        )
    logger.info(
        "AD lookup: SUCCESS for emp_id=%r - %d candidate(s), emp_name=%r manager_name=%r",
        emp_id, len(candidates), candidates[0].get("emp_name"), candidates[0].get("manager_name"),
    )
    return candidates


def _run_webview_helper():
    import webview

    logger.info("AD helper: subprocess started (pid=%s)", os.getpid())

    # Single-file app: settings live in the CONFIG dict at the top of this
    # file, not a separate config.json - same values, just one less file.
    # This subprocess is launched fresh (re-imports this module) and never
    # sees a running WizardApp's self.config_data, so any admin_settings.json
    # override (e.g. a changed AD Lookup URL) must be re-applied here
    # independently, or it would silently keep using the CONFIG default.
    config = CONFIG
    _apply_admin_settings_overrides(config)
    mode = config.get("ad_lookup_mode", "ssrs_report")
    profile_dir = os.path.join(os.path.expanduser("~"), ".it_asset_form_webview_profile")

    if mode == "ssrs_report":
        ssrs_cfg = config.get("ad_lookup_ssrs", {})
        search_url = ssrs_cfg.get("search_page_url", "")
        button_text = ssrs_cfg.get("view_report_button_text", "View Report")
        header_keyword = ssrs_cfg.get("results_header_keyword", "EmployeeID")
        name_column = ssrs_cfg.get("employee_name_column", "Employee Name")
        ready_timeout = ssrs_cfg.get("results_ready_timeout_seconds", 25)

        if not search_url or search_url.startswith("REPLACE_ME"):
            _send({
                "cmd": "fatal_error",
                "error": (
                    "CONFIG -> ad_lookup_ssrs.search_page_url must be set "
                    "to your real 'Active Directory Search' report URL. "
                    "See README.md."
                ),
            })
            return
    else:
        browser_cfg = config.get("ad_lookup_browser", {})
        search_url = browser_cfg.get("search_page_url", "")
        input_sel = browser_cfg.get("search_input_selector", "")
        submit_sel = browser_cfg.get("search_submit_selector", "")
        use_enter = browser_cfg.get("use_enter_key_to_submit", False)
        results_ready_sel = browser_cfg.get("results_ready_selector", "")
        multi_row_sel = browser_cfg.get("multi_result_row_selector", "")
        ready_timeout = browser_cfg.get("results_ready_timeout_seconds", 20)

        if not search_url or not input_sel:
            _send({
                "cmd": "fatal_error",
                "error": (
                    "CONFIG -> ad_lookup_browser.search_page_url and "
                    "search_input_selector must be set before this can "
                    "work. See README.md."
                ),
            })
            return

    window = webview.create_window(
        "Employee Lookup - log in here if prompted, then leave this window open",
        search_url,
        width=1000,
        height=800,
    )

    page_ready = threading.Event()
    window.events.loaded += lambda: page_ready.set()

    if mode == "ssrs_report":
        def perform_search(emp_id):
            return _perform_ssrs_search(
                window, emp_id, button_text, header_keyword, name_column, ready_timeout
            )
    else:
        def perform_search(emp_id):
            safe_id = json.dumps(emp_id)  # safe JS string literal

            fill_js = f"""
            (function() {{
                var el = document.querySelector({json.dumps(input_sel)});
                if (!el) return "NO_INPUT";
                el.focus();
                el.value = {safe_id};
                el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                return "OK";
            }})();
            """
            if window.evaluate_js(fill_js) != "OK":
                raise RuntimeError(
                    f"Could not find the search box (selector '{input_sel}'). "
                    "Update ad_lookup_browser.search_input_selector in CONFIG."
                )

            if use_enter or not submit_sel:
                enter_js = f"""
                (function() {{
                    var el = document.querySelector({json.dumps(input_sel)});
                    if (!el) return "NO_INPUT";
                    el.dispatchEvent(new KeyboardEvent('keydown', {{
                        key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
                    }}));
                    return "OK";
                }})();
                """
                window.evaluate_js(enter_js)
            else:
                click_js = f"""
                (function() {{
                    var btn = document.querySelector({json.dumps(submit_sel)});
                    if (!btn) return "NO_BUTTON";
                    btn.click();
                    return "OK";
                }})();
                """
                if window.evaluate_js(click_js) != "OK":
                    raise RuntimeError(
                        f"Could not find the search/submit button (selector "
                        f"'{submit_sel}'). Update ad_lookup_browser."
                        "search_submit_selector in CONFIG, or set "
                        "use_enter_key_to_submit to true."
                    )

            # Wait for results to render.
            if results_ready_sel:
                check_js = f"return !!document.querySelector({json.dumps(results_ready_sel)});"
                deadline = time.time() + ready_timeout
                found = False
                while time.time() < deadline:
                    if window.evaluate_js(check_js):
                        found = True
                        break
                    time.sleep(0.3)
                if not found:
                    raise RuntimeError(
                        "Timed out waiting for search results to appear "
                        f"(looked for '{results_ready_sel}'). The employee ID "
                        "may not exist, or results_ready_selector needs "
                        "adjusting in CONFIG."
                    )
            else:
                time.sleep(min(3, ready_timeout))

            # Extract candidate(s).
            if multi_row_sel:
                rows_js = f"""
                (function() {{
                    var rows = document.querySelectorAll({json.dumps(multi_row_sel)});
                    return Array.prototype.map.call(rows, function(r) {{ return r.outerHTML; }});
                }})();
                """
                row_htmls = window.evaluate_js(rows_js) or []
            else:
                row_htmls = [window.evaluate_js("document.documentElement.outerHTML")]

            candidates = []
            for html in row_htmls:
                lines = _html_to_lines(html)
                fields = _extract_fields(lines)
                if any(fields.values()):
                    candidates.append(fields)

            if not candidates:
                raise RuntimeError(
                    "Search ran, but no employee fields could be parsed out of "
                    "the results. Check FIELD_LABEL_MAP (near the top of this "
                    "file) against your real results' label text."
                )
            return candidates

    def stdin_reader():
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            try:
                req = json.loads(line)
            except Exception:
                continue

            if req.get("cmd") == "shutdown":
                try:
                    window.destroy()
                except Exception:
                    pass
                return

            request_id = req.get("request_id")
            emp_id = (req.get("emp_id") or "").strip()
            logger.info("AD helper: received lookup request_id=%r emp_id=%r", request_id, emp_id)

            # The window may have been auto-hidden after the previous
            # successful lookup (see below) - bring it back so the user can
            # see this new search happen (and log in again if a session
            # expired in the meantime).
            try:
                window.show()
            except Exception:
                logger.exception("AD helper: window.show() failed (request_id=%r)", request_id)

            if not emp_id:
                logger.warning("AD helper: empty employee ID in request_id=%r", request_id)
                _send({"request_id": request_id, "error": "Empty employee ID."})
                continue

            if not page_ready.wait(timeout=180):
                logger.error("AD helper: browser window never finished loading (request_id=%r)", request_id)
                _send({
                    "request_id": request_id,
                    "error": "Browser window never finished loading (login taking too long?).",
                })
                continue

            try:
                candidates = perform_search(emp_id)
                logger.info(
                    "AD helper: sending %d candidate(s) back for request_id=%r",
                    len(candidates), request_id,
                )
                _send({"request_id": request_id, "candidates": candidates})
                # Lookup succeeded and the app now has the data it needs -
                # tuck the lookup window out of the way automatically instead
                # of leaving it sitting on top of the main form. It stays
                # alive in the background (not destroyed) so the next lookup
                # is instant and window.show() above can bring it right back.
                if candidates:
                    try:
                        window.hide()
                        logger.info("AD helper: auto-hid lookup window after successful result (request_id=%r)", request_id)
                    except Exception:
                        logger.exception("AD helper: window.hide() failed (request_id=%r)", request_id)
            except Exception as exc:
                logger.exception("AD helper: lookup failed for request_id=%r emp_id=%r", request_id, emp_id)
                _send({"request_id": request_id, "error": str(exc)})

    threading.Thread(target=stdin_reader, daemon=True).start()

    # gui="edgechromium" forces the WebView2/Edge backend on Windows.
    # private_mode=False + storage_path keeps the session/cookies alive on
    # disk too, as a bonus, in case the app is restarted later.
    webview.start(gui="edgechromium", private_mode=False, storage_path=profile_dir)


# ============================================================================
# BULK BATCH QUEUE - reads the Excel import file
# ============================================================================

class BulkQueueError(Exception):
    pass


COLUMN_ALIASES = {
    "employee_id": ["employee id", "employeeid", "empid", "emp id", "id"],
    "type": ["type", "title", "submissiontype", "submission type", "category"],
    "employee_name": ["employee name", "employeename", "name", "emp name"],
    "manager_name": ["manager name", "managername", "manager", "reporting manager"],
    "contact_number": ["contact number", "contactnumber", "phone", "mobile"],
    "business_unit": ["business unit", "businessunit", "bu"],
    "last_working_date": ["last working date", "lastworkingdate", "exit date", "exitdate", "lwd date"],
    "laptop_serial_number": [
        "laptop serial number", "laptop serial", "laptopserialnumber", "device serial number",
        "device serial", "serial number", "serialnumber", "serial no", "serial",
    ],
}


def _normalize(header):
    return str(header).strip().lower() if header is not None else ""


def _map_columns(header_row):
    """Returns {field_name: column_index} for whichever recognized columns
    are present in this header row."""
    normalized = [_normalize(h) for h in header_row]
    mapping = {}
    for field, aliases in COLUMN_ALIASES.items():
        for idx, h in enumerate(normalized):
            if h in aliases:
                mapping[field] = idx
                break
    return mapping


def resolve_submission_type(raw_type, type_aliases):
    """Maps a raw 'Type' cell value (e.g. 'Exit', 'NewHire') to one of the
    app's actual submission types, using config.json's bulk_type_aliases.
    Returns (resolved_type, was_recognized)."""
    key = str(raw_type).strip().lower() if raw_type else ""
    for canonical, aliases in type_aliases.items():
        if key == canonical.lower() or key in [a.lower() for a in aliases]:
            return canonical, True
    return str(raw_type).strip(), False


def load_queue(xlsx_path, type_aliases, sheet_name=None):
    """
    Returns (queue, warnings):
      queue    - list of dicts, one per row, with keys: employee_id, type,
                 type_recognized (bool), employee_name, manager_name,
                 contact_number, business_unit, laptop_serial_number,
                 row_number (for error messages), status ("Pending")
      warnings - list of human-readable strings for rows that were
                 skipped or had an unrecognized Type - shown to the user
                 before they start processing, so nothing is silently lost.
    """
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    except Exception as exc:
        raise BulkQueueError(f"Could not open '{xlsx_path}': {exc}")

    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise BulkQueueError(f"'{ws.title}' is empty.")

    header_row = rows[0]
    mapping = _map_columns(header_row)

    if "employee_id" not in mapping:
        raise BulkQueueError(
            "Could not find an Employee ID column. Expected a header like "
            "'Employee ID', 'EmployeeID', 'EmpID', or 'ID'. "
            f"Found headers: {[h for h in header_row if h]}"
        )
    if "type" not in mapping:
        raise BulkQueueError(
            "Could not find a Type column. Expected a header like 'Type' or "
            f"'Title'. Found headers: {[h for h in header_row if h]}"
        )

    queue = []
    warnings = []

    def cell(row, field):
        idx = mapping.get(field)
        if idx is None or idx >= len(row):
            return None
        val = row[idx]
        return str(val).strip() if val is not None else None

    for i, row in enumerate(rows[1:], start=2):  # start=2: row 1 is the header
        emp_id = cell(row, "employee_id")
        raw_type = cell(row, "type")

        if not emp_id and not raw_type:
            continue  # blank row - skip silently, this is normal at the end of a sheet

        if not emp_id:
            warnings.append(f"Row {i}: no Employee ID - skipped.")
            continue
        if not raw_type:
            warnings.append(f"Row {i} (Employee ID {emp_id}): no Type - skipped.")
            continue

        resolved_type, recognized = resolve_submission_type(raw_type, type_aliases)
        if not recognized:
            warnings.append(
                f"Row {i} (Employee ID {emp_id}): Type '{raw_type}' isn't recognized - "
                "kept as entered, but the app won't know which field pattern or email "
                "rule to apply until you add it to config.json -> bulk_type_aliases, "
                "or you can fix it manually for this person on the day."
            )

        queue.append({
            "employee_id": emp_id,
            "type": resolved_type,
            "type_recognized": recognized,
            "employee_name": cell(row, "employee_name"),
            "manager_name": cell(row, "manager_name"),
            "contact_number": cell(row, "contact_number"),
            "business_unit": cell(row, "business_unit"),
            "last_working_date": cell(row, "last_working_date"),
            "laptop_serial_number": cell(row, "laptop_serial_number"),
            "row_number": i,
            "status": "Pending",
        })

    if not queue:
        raise BulkQueueError("No usable rows found - check the warnings above.")

    return queue, warnings


def create_batch_template_xlsx(path, config_data):
    """Writes a starter Excel file for bulk import: a 'Batch' sheet with
    the recognized column headers plus a few example rows, and an
    'Instructions' sheet explaining each column and listing the valid
    Type values. Meant to be downloaded once by anyone unfamiliar with
    the expected format, filled in, and fed back into 'Import Batch from
    Excel...' on the Bulk Batch tab."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Batch"

    headers = ["Employee ID", "Type", "Employee Name", "Manager Name",
               "Contact Number", "Business Unit", "Last Working Date", "Laptop Serial Number"]
    ws.append(headers)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = header_fill
        c.alignment = Alignment(horizontal="left")

    example_rows = [
        ["900012345", "New Hire", "", "", "9876543210", "", "", "LAP-2026-0091"],
        ["900054321", "LWD", "Jane Smith", "Priya Kumar", "9123456780", "", "30-Sep-2026", "LAP-2023-0417"],
        ["900099887", "Break Fix", "", "", "", "My Location", "", ""],
    ]
    for row in example_rows:
        ws.append(row)

    widths = [16, 16, 22, 22, 16, 18, 18, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    instr = wb.create_sheet("Instructions")
    lines = [
        ("How to use this file", True),
        ("1. Fill in one row per person on the 'Batch' sheet - keep the header row as-is.", False),
        ("2. Only 'Employee ID' and 'Type' are required. Everything else is optional.", False),
        ("3. Save the file, then use 'Import Batch from Excel...' on the Bulk Batch tab.", False),
        ("", False),
        ("Column reference", True),
        ("Employee ID - required. The person's ID, exactly as AD/your HR system has it.", False),
        ("Type - required. One of the exact values listed below (not case-sensitive; common", False),
        ("  synonyms like 'exit' or 'joiner' are also recognized).", False),
        ("Employee Name / Manager Name - optional. Leave BOTH blank to have the app look these", False),
        ("  up automatically via AD when this person is processed. Fill them in to skip that", False),
        ("  lookup (recommended for exits/LWD, where this is usually already known).", False),
        ("Contact Number - optional. Can be filled in later in the app if left blank.", False),
        ("Business Unit - optional. Must exactly match a Business Unit name already added in the", False),
        ("  app. Leave blank to use whichever Business Unit is selected for this batch in the app.", False),
        ("Last Working Date - optional. Only used for exit/LWD-style types.", False),
        ("Laptop Serial Number - optional. Fill it in wherever you already know it and the app will", False),
        ("  pre-fill the matching serial number field for that person (the 'new' serial for New Hire,", False),
        ("  the 'current/returning' serial for everything else) once you open them to process. Leave", False),
        ("  blank to type it in manually when you get to that person.", False),
        ("", False),
        ("Valid Type values", True),
    ]
    for label in config_data.get("submission_type_row1", []) + config_data.get("submission_type_row2", []):
        lines.append((f"  - {label}", False))

    for r, (text, is_heading) in enumerate(lines, start=1):
        cell = instr.cell(row=r, column=1, value=text)
        if is_heading:
            cell.font = Font(bold=True, size=12)
    instr.column_dimensions["A"].width = 95

    wb.save(path)


# ============================================================================
# DOCX CONTROLS - turns a raw uploaded BU Word doc into a fillable one
# ============================================================================

CHECKBOX_TAGS = [
    "NewHire", "BreakFix", "MixedBuild", "AdditionalLaptop", "LaptopReturn",
    "ContractorLWD", "LWD", "SiteTransfer", "RCP",
    "LaptopBag", "Charger", "Mouse", "Headset", "NetworkCable", "Multiport",
]
BLANK_TAGS = ["Date", "EmpID", "EmpName", "ContactNumber", "ManagerName", "LastWorkingDate"]
CELL_TAGS = [
    ("Current Device Serial Number", "CurrentSerial", "text"),
    ("New Device Serial Number", "NewSerial", "text"),
    ("Asset Pending for Submission", "AssetPending", "text"),
    ("Signature Of Employee", "EmployeeSignature", "picture"),
    ("Asset Receiver Signature", "AssetReceiverSignature", "picture"),
]
ASSET_TEXT_TAGS = ["Others", "CurrentSerial", "NewSerial", "AssetPending"]
SIG_REL_ID = "rIdSigPlaceholder"

# Every field this app fills into the BU's Word template. Kept here (not
# just inline) so both insert_controls() and fill_docx() agree on the
# exact same list.
REQUIRED_TEXT_FIELDS = BLANK_TAGS + CHECKBOX_TAGS + ASSET_TEXT_TAGS
SIGNATURE_TAGS = ["EmployeeSignature", "AssetReceiverSignature"]


class ControlInsertionError(Exception):
    """Raised when the uploaded document doesn't match the expected
    layout closely enough to insert controls automatically."""
    pass


class FillDocxError(Exception):
    """Raised when a previously-inserted control can't be found while
    filling a document - almost always means the BU template is stale
    (was added before a code change) and should be re-added."""
    pass


def _new_id_counter():
    n = [900000]
    def next_id():
        n[0] += 1
        return n[0]
    return next_id


def _text_sdt(tag, inner_runs, next_id):
    return (
        f'<w:sdt><w:sdtPr>'
        f'<w:alias w:val="{tag}"/><w:tag w:val="{tag}"/>'
        f'<w:id w:val="{next_id()}"/>'
        f'<w:lock w:val="sdtLocked"/>'
        f'<w:text/>'
        f'</w:sdtPr><w:sdtContent>{inner_runs}</w:sdtContent></w:sdt>'
    )


def _picture_sdt(tag, rel_id, next_id):
    drawing = (
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="1800000" cy="600000"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{next_id() % 100000}" name="{tag}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="0" name="{tag}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rel_id}"/>'
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="600000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    )
    return (
        f'<w:sdt><w:sdtPr>'
        f'<w:alias w:val="{tag}"/><w:tag w:val="{tag}"/>'
        f'<w:id w:val="{next_id()}"/>'
        f'<w:picture/>'
        f'</w:sdtPr><w:sdtContent>{drawing}</w:sdtContent></w:sdt>'
    )


def _insert_controls_into_xml(xml, next_id):
    run_pattern = re.compile(r"<w:r(?: [^>]*)?>(?:(?!</w:r>).)*?</w:r>", re.S)

    # ---- checkboxes ----
    cb_iter = iter(CHECKBOX_TAGS)
    matched_checkboxes = []

    def wrap_checkbox(m):
        run = m.group(0)
        t = re.search(r"<w:t[^>]*>([^<]*)</w:t>", run)
        if not t or "\u2610" not in t.group(1):
            return run
        try:
            tag = next(cb_iter)
        except StopIteration:
            return run
        matched_checkboxes.append(tag)
        return _text_sdt(tag, run, next_id)

    xml = run_pattern.sub(wrap_checkbox, xml)
    if len(matched_checkboxes) < len(CHECKBOX_TAGS):
        missing = CHECKBOX_TAGS[len(matched_checkboxes):]
        raise ControlInsertionError(
            f"Expected {len(CHECKBOX_TAGS)} checkboxes (\u2610 characters) in the document, "
            f"but only found {len(matched_checkboxes)}. Missing/unmatched: {', '.join(missing)}. "
            "This document's checkbox list may not match the standard form."
        )

    # ---- underscore blanks ----
    runs = list(run_pattern.finditer(xml))

    def run_text(run_xml):
        m = re.search(r"<w:t[^>]*>([^<]*)</w:t>", run_xml)
        return m.group(1) if m else ""

    groups, current = [], []
    for m in runs:
        txt = run_text(m.group(0))
        is_blank = bool(txt) and re.fullmatch(r"[_\s]+", txt) and "_" in txt
        if is_blank:
            if current and current[-1].end() == m.start():
                current.append(m)
            else:
                if current:
                    groups.append(current)
                current = [m]
        else:
            if current:
                groups.append(current)
                current = []
    if current:
        groups.append(current)

    if len(groups) < len(BLANK_TAGS):
        raise ControlInsertionError(
            f"Expected {len(BLANK_TAGS)} underline blank fields (Date, Emp ID, Emp Name, "
            f"Contact Number, Manager Name, Last Working Date) but only found {len(groups)} "
            "underscore-blank runs in the document."
        )

    replacements = []
    for gi, group in enumerate(groups[: len(BLANK_TAGS)]):
        tag = BLANK_TAGS[gi]
        start, end = group[0].start(), group[-1].end()
        first = group[0].group(0)
        rpr = re.search(r"<w:rPr>.*?</w:rPr>", first, re.S)
        rpr_xml = rpr.group(0) if rpr else ""
        placeholder_run = f'<w:r>{rpr_xml}<w:t xml:space="preserve">     </w:t></w:r>'
        replacements.append((start, end, _text_sdt(tag, placeholder_run, next_id)))

    for start, end, new_xml in sorted(replacements, reverse=True):
        xml = xml[:start] + new_xml + xml[end:]

    # ---- "Others:" field ----
    others_run = re.search(
        r'<w:r(?: [^>]*)?>((?:(?!</w:r>).)*?)<w:t[^>]*>Others:[^<]*</w:t></w:r>', xml, re.S
    )
    if not others_run:
        raise ControlInsertionError('Could not find the "Others:" line in the assets table.')
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", others_run.group(0), re.S)
    rpr_xml = rpr.group(0) if rpr else ""
    label_run = f'<w:r>{rpr_xml}<w:t xml:space="preserve">Others: </w:t></w:r>'
    value_run = f'<w:r>{rpr_xml}<w:t xml:space="preserve">                              </w:t></w:r>'
    xml = (
        xml[: others_run.start()]
        + label_run
        + _text_sdt("Others", value_run, next_id)
        + xml[others_run.end():]
    )

    # ---- table cells (by preceding label) ----
    for label, tag, kind in CELL_TAGS:
        li = xml.find(label)
        if li == -1:
            raise ControlInsertionError(f'Could not find the "{label}" label in the document.')
        cell_end = xml.find("</w:tc>", li)
        next_cell_start = xml.find("<w:tc>", cell_end)
        next_cell_end = xml.find("</w:tc>", next_cell_start)
        if next_cell_start == -1 or next_cell_end == -1:
            raise ControlInsertionError(f'Could not find the value cell after "{label}".')
        cell_xml = xml[next_cell_start:next_cell_end]

        p_match = re.search(r"<w:p\b[^>]*>(?:(?!</w:p>).)*?</w:p>", cell_xml, re.S)
        if not p_match:
            raise ControlInsertionError(f'No paragraph found in the cell after "{label}".')

        if kind == "text":
            content = _text_sdt(tag, '<w:r><w:t xml:space="preserve">     </w:t></w:r>', next_id)
        else:
            content = _picture_sdt(tag, SIG_REL_ID, next_id)

        p_xml = p_match.group(0)
        new_p = p_xml[: p_xml.rfind("</w:p>")] + content + "</w:p>"
        new_cell = cell_xml[: p_match.start()] + new_p + cell_xml[p_match.end():]
        xml = xml[:next_cell_start] + new_cell + xml[next_cell_end:]

    return xml


def insert_controls(source_docx_path, output_docx_path, placeholder_image_path):
    """
    Reads source_docx_path, inserts all 27 content controls, writes the
    result to output_docx_path. placeholder_image_path is a small PNG used
    to satisfy the picture-control requirement that it hold *some* image
    (fill_docx() replaces it with the real signature later).

    Raises ControlInsertionError with a specific, actionable message if the
    document's layout doesn't match closely enough - this is surfaced
    directly to the person uploading the template.
    """
    logger.info("insert_controls: starting for %r -> %r", source_docx_path, output_docx_path)
    with tempfile.TemporaryDirectory() as tmp:
        unpack_dir = os.path.join(tmp, "unpacked")
        with zipfile.ZipFile(source_docx_path) as z:
            z.extractall(unpack_dir)

        doc_xml_path = os.path.join(unpack_dir, "word", "document.xml")
        if not os.path.exists(doc_xml_path):
            logger.error("insert_controls: %r is not a valid .docx (no word/document.xml)", source_docx_path)
            raise ControlInsertionError("This doesn't look like a valid .docx file.")

        with open(doc_xml_path, encoding="utf-8") as f:
            xml = f.read()

        next_id = _new_id_counter()
        xml = _insert_controls_into_xml(xml, next_id)

        # Ensure required namespaces for the picture controls are declared.
        m = re.search(r"<w:document[^>]*>", xml)
        root = m.group(0)
        new_root = root
        if "xmlns:a=" not in new_root:
            new_root = new_root[:-1] + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        if "xmlns:pic=" not in new_root:
            new_root = new_root[:-1] + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        xml = xml.replace(root, new_root, 1)

        with open(doc_xml_path, "w", encoding="utf-8") as f:
            f.write(xml)

        # Add the placeholder image + relationship for the picture controls.
        media_dir = os.path.join(unpack_dir, "word", "media")
        os.makedirs(media_dir, exist_ok=True)
        shutil.copy2(placeholder_image_path, os.path.join(media_dir, "sig_placeholder.png"))

        rels_path = os.path.join(unpack_dir, "word", "_rels", "document.xml.rels")
        with open(rels_path, encoding="utf-8") as f:
            rels = f.read()
        if SIG_REL_ID not in rels:
            new_rel = (
                f'<Relationship Id="{SIG_REL_ID}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                'Target="media/sig_placeholder.png"/>'
            )
            rels = rels.replace("</Relationships>", new_rel + "</Relationships>")
            with open(rels_path, "w", encoding="utf-8") as f:
                f.write(rels)

        # Repack.
        if os.path.exists(output_docx_path):
            os.remove(output_docx_path)
        with zipfile.ZipFile(output_docx_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root_dir, _, files in os.walk(unpack_dir):
                for file in files:
                    full = os.path.join(root_dir, file)
                    arcname = os.path.relpath(full, unpack_dir)
                    zf.write(full, arcname)

    logger.info("insert_controls: SUCCESS -> %r", output_docx_path)
    return output_docx_path


def _default_placeholder_image_path():
    """A tiny transparent PNG bundled/generated next to the app, used only
    to satisfy the picture-control requirement of holding *some* image
    when a BU template is first added - fill_docx() always replaces it
    with the real captured signature before anyone sees the output."""
    path = os.path.join(_app_dir(), "assets", "_sig_placeholder.png")
    if not os.path.exists(path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        Image.new("RGBA", (400, 120), (255, 255, 255, 0)).save(path)
    return path


def _xml_escape(text):
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _extract_rpr(inner_xml):
    m = re.search(r"<w:rPr>.*?</w:rPr>", inner_xml, re.S)
    return m.group(0) if m else ""


def _sdt_pattern(tag):
    return re.compile(
        r'<w:sdt><w:sdtPr>.*?<w:tag w:val="' + re.escape(tag) + r'"/>.*?</w:sdtPr>'
        r'<w:sdtContent>(.*?)</w:sdtContent></w:sdt>',
        re.S,
    )


def _replace_sdt_content(xml, tag, new_inner_xml, required=True):
    pattern = re.compile(
        r'(<w:sdt><w:sdtPr>.*?<w:tag w:val="' + re.escape(tag) + r'"/>.*?</w:sdtPr><w:sdtContent>)'
        r'.*?(</w:sdtContent></w:sdt>)',
        re.S,
    )
    m = pattern.search(xml)
    if not m:
        if required:
            raise FillDocxError(
                f"Could not find the '{tag}' field in this template - it may be out of "
                "date. Re-add this Business Unit's template via 'Add New BU Template...'."
            )
        return xml
    return xml[: m.start()] + m.group(1) + new_inner_xml + m.group(2) + xml[m.end():]


def fill_docx(fillable_docx_path, output_docx_path, field_values, checked_tags, signature_images):
    """
    Fills every text/checkbox/signature-picture control previously inserted
    by insert_controls() into a copy of the BU's Word template.

    field_values: {tag: text} - any subset of BLANK_TAGS/ASSET_TEXT_TAGS;
        tags not on this particular form's variant are silently skipped.
    checked_tags: set of CHECKBOX_TAGS that should render checked - every
        other checkbox tag is explicitly rendered unchecked, so a previous
        run's checked boxes never carry over.
    signature_images: {"EmployeeSignature": path, "AssetReceiverSignature":
        path} - the real captured signature images, dropped in as the
        visible picture inside each signature control. This is what makes
        the exported PDF visually show both signatures before the
        cryptographic signing step (sign_pdf_with_signatures) runs on it.
    """
    logger.info(
        "fill_docx: starting for %r -> %r (fields=%d, checked_tags=%r, signatures=%r)",
        fillable_docx_path, output_docx_path, len(field_values), sorted(checked_tags),
        {k: bool(v) for k, v in signature_images.items()},
    )
    with tempfile.TemporaryDirectory() as tmp:
        unpack_dir = os.path.join(tmp, "unpacked")
        with zipfile.ZipFile(fillable_docx_path) as z:
            z.extractall(unpack_dir)

        doc_xml_path = os.path.join(unpack_dir, "word", "document.xml")
        with open(doc_xml_path, encoding="utf-8") as f:
            xml = f.read()

        # ---- text fields ----
        for tag, value in field_values.items():
            m = _sdt_pattern(tag).search(xml)
            if not m:
                continue  # this tag isn't on this particular BU form variant - fine to skip
            rpr = _extract_rpr(m.group(1))
            safe_value = _xml_escape(value) if value else " "
            new_inner = f'<w:r>{rpr}<w:t xml:space="preserve">{safe_value}</w:t></w:r>'
            xml = _replace_sdt_content(xml, tag, new_inner)

        # ---- checkboxes ----
        for tag in CHECKBOX_TAGS:
            m = _sdt_pattern(tag).search(xml)
            if not m:
                continue
            glyph = "\u2611" if tag in checked_tags else "\u2610"
            new_inner = re.sub(r"(<w:t[^>]*>)[^<]*(</w:t>)", rf"\g<1>{glyph}\g<2>", m.group(1), count=1)
            xml = _replace_sdt_content(xml, tag, new_inner)

        # ---- signature pictures ----
        media_dir = os.path.join(unpack_dir, "word", "media")
        os.makedirs(media_dir, exist_ok=True)
        rels_path = os.path.join(unpack_dir, "word", "_rels", "document.xml.rels")
        with open(rels_path, encoding="utf-8") as f:
            rels = f.read()

        for tag in SIGNATURE_TAGS:
            image_path = signature_images.get(tag)
            if not image_path:
                continue
            m = _sdt_pattern(tag).search(xml)
            if not m:
                continue
            rel_id = f"rId{tag}Signed"
            media_name = f"{tag}_signed.png"
            shutil.copy2(image_path, os.path.join(media_dir, media_name))
            if rel_id not in rels:
                new_rel = (
                    f'<Relationship Id="{rel_id}" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                    f'Target="media/{media_name}"/>'
                )
                rels = rels.replace("</Relationships>", new_rel + "</Relationships>")
            new_inner = re.sub(r'r:embed="[^"]*"', f'r:embed="{rel_id}"', m.group(1), count=1)
            xml = _replace_sdt_content(xml, tag, new_inner)

        with open(rels_path, "w", encoding="utf-8") as f:
            f.write(rels)
        with open(doc_xml_path, "w", encoding="utf-8") as f:
            f.write(xml)

        if os.path.exists(output_docx_path):
            os.remove(output_docx_path)
        with zipfile.ZipFile(output_docx_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root_dir, _, files in os.walk(unpack_dir):
                for file in files:
                    full = os.path.join(root_dir, file)
                    arcname = os.path.relpath(full, unpack_dir)
                    zf.write(full, arcname)

    logger.info("fill_docx: SUCCESS -> %r", output_docx_path)
    return output_docx_path


# ============================================================================
# BUSINESS UNIT TEMPLATES
# ============================================================================

def _manifest_path(templates_folder):
    return os.path.join(templates_folder, "manifest.json")


def _load_manifest(templates_folder):
    path = _manifest_path(templates_folder)
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _save_manifest(templates_folder, manifest):
    os.makedirs(templates_folder, exist_ok=True)
    with open(_manifest_path(templates_folder), "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)


def list_templates(templates_folder):
    """Returns the list of saved BU templates: [{"id", "name", "fillable_path"}, ...]"""
    return _load_manifest(templates_folder)


def get_template(templates_folder, bu_id):
    for entry in _load_manifest(templates_folder):
        if entry["id"] == bu_id:
            return entry
    return None


def delete_template(templates_folder, bu_id):
    manifest = _load_manifest(templates_folder)
    manifest = [e for e in manifest if e["id"] != bu_id]
    _save_manifest(templates_folder, manifest)
    bu_dir = os.path.join(templates_folder, bu_id)
    if os.path.isdir(bu_dir):
        shutil.rmtree(bu_dir, ignore_errors=True)


def add_template(templates_folder, bu_name, source_docx_path):
    """
    Adds a BU's base Word (.docx) form: automatically inserts every
    fillable field/checkbox/signature control this app needs into a copy
    of it (see insert_controls above) and saves that alongside the
    original. Returns the new manifest entry.

    Raises ControlInsertionError with a SPECIFIC message (naming exactly
    which part of the expected layout - a checkbox, an underline blank, a
    table label - couldn't be found) if the document's layout doesn't
    match closely enough. Show that message directly to the person
    uploading it, since it tells them exactly what to check in the Word
    doc before re-uploading.
    """
    bu_id = uuid.uuid4().hex[:10]
    bu_dir = os.path.join(templates_folder, bu_id)
    os.makedirs(bu_dir, exist_ok=True)

    original_path = os.path.join(bu_dir, "original.docx")
    shutil.copy2(source_docx_path, original_path)

    fillable_path = os.path.join(bu_dir, "fillable.docx")
    try:
        insert_controls(original_path, fillable_path, _default_placeholder_image_path())
    except ControlInsertionError:
        shutil.rmtree(bu_dir, ignore_errors=True)
        raise

    entry = {
        "id": bu_id,
        "name": bu_name.strip(),
        "fillable_path": fillable_path,
    }

    manifest = _load_manifest(templates_folder)
    manifest.append(entry)
    _save_manifest(templates_folder, manifest)
    return entry


# ============================================================================
# WORD -> PDF CONVERSION (uses a real, locally installed Microsoft Word)
# ============================================================================

class WordConversionError(Exception):
    pass


def convert_docx_to_pdf_via_word(docx_path, pdf_path):
    """
    Uses Microsoft Word itself (via COM automation) to export the filled
    document to PDF - the same "Save As -> PDF" Word already offers, just
    scripted, so the exported page looks exactly like the real form. Needs
    Word installed on this machine. Adobe Acrobat is never involved.
    """
    logger.info("Word export: starting for %r -> %r", docx_path, pdf_path)
    start_time = time.time()
    try:
        import win32com.client as win32
    except ImportError as exc:
        logger.error("Word export: pywin32 isn't installed")
        raise WordConversionError(
            "pywin32 isn't installed. Run: pip install pywin32"
        ) from exc

    docx_abspath = os.path.abspath(docx_path)
    pdf_abspath = os.path.abspath(pdf_path)
    os.makedirs(os.path.dirname(pdf_abspath) or ".", exist_ok=True)
    if os.path.exists(pdf_abspath):
        os.remove(pdf_abspath)

    word = None
    doc = None
    try:
        logger.debug("Word export: launching Word.Application (DispatchEx)")
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0
        except Exception as exc:
            logger.debug("Word export: could not set DisplayAlerts=0 (%s)", exc)
        logger.debug("Word export: opening %r", docx_abspath)
        doc = word.Documents.Open(docx_abspath, ReadOnly=True)
        logger.debug("Word export: saving as PDF -> %r", pdf_abspath)
        doc.SaveAs(pdf_abspath, FileFormat=17)  # wdFormatPDF
    except Exception as exc:
        logger.exception("Word export: FAILED after %.1fs", time.time() - start_time)
        raise WordConversionError(
            f"Could not convert the filled document to PDF using Microsoft Word: {exc}\n"
            "Make sure Microsoft Word is installed on this machine."
        ) from exc
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            logger.debug("Word export: doc.Close() failed (non-fatal)", exc_info=True)
        try:
            if word is not None:
                word.Quit()
        except Exception:
            logger.debug("Word export: word.Quit() failed (non-fatal)", exc_info=True)

    if not os.path.exists(pdf_abspath):
        logger.error(
            "Word export: Word reported success but %r does not exist (elapsed %.1fs)",
            pdf_abspath, time.time() - start_time,
        )
        raise WordConversionError(
            "Word reported success but no PDF file was produced - try again, "
            "or check no other Word window is blocking a dialog."
        )
    logger.info("Word export: SUCCESS -> %r (%.1fs)", pdf_abspath, time.time() - start_time)
    return pdf_abspath


# ============================================================================
# SIGNATURE CAPTURE - Topaz pad (if enabled/available) or on-screen canvas
# ============================================================================

class TopazNotAvailable(Exception):
    """Raised when a Topaz pad was requested but couldn't be used (SDK
    missing, pad not connected, nothing signed) - SignatureService catches
    this and falls back to the on-screen canvas automatically."""
    pass


class CanvasSigner:
    """On-screen mouse/touch signature capture - always available, and
    what's used automatically whenever a physical Topaz pad isn't
    connected/enabled (including for testing on a machine with no pad)."""

    WIDTH, HEIGHT = 500, 200

    def capture(self, parent, title="Sign here"):
        result = {"path": None}
        win = tk.Toplevel(parent)
        win.title(title)
        win.grab_set()
        win.resizable(False, False)

        ttk.Label(win, text=title, font=("Segoe UI", 10, "bold")).pack(padx=12, pady=(12, 4))
        canvas = tk.Canvas(win, width=self.WIDTH, height=self.HEIGHT, bg="white",
                            cursor="pencil", highlightthickness=1, highlightbackground="#999")
        canvas.pack(padx=12, pady=(0, 4))
        ttk.Label(win, text="Draw your signature above with the mouse, then click Accept.",
                  foreground="#555").pack()

        strokes = []
        current_stroke = []

        def on_press(event):
            current_stroke.clear()
            current_stroke.append((event.x, event.y))

        def on_drag(event):
            if current_stroke:
                x0, y0 = current_stroke[-1]
                canvas.create_line(x0, y0, event.x, event.y, width=2, fill="black",
                                    capstyle=tk.ROUND, smooth=True)
            current_stroke.append((event.x, event.y))

        def on_release(_event):
            if len(current_stroke) > 1:
                strokes.append(list(current_stroke))
            current_stroke.clear()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)

        def clear():
            canvas.delete("all")
            strokes.clear()

        def accept():
            if not strokes:
                messagebox.showwarning("Not signed", "Please draw a signature before clicking Accept.")
                return
            img = Image.new("RGB", (self.WIDTH, self.HEIGHT), "white")
            draw = ImageDraw.Draw(img)
            for stroke in strokes:
                if len(stroke) > 1:
                    draw.line(stroke, fill="black", width=2, joint="curve")
            tmp_path = os.path.join(tempfile.gettempdir(), f"sig_{uuid.uuid4().hex}.png")
            img.save(tmp_path)
            result["path"] = tmp_path
            win.destroy()

        def cancel():
            win.destroy()

        btn_row = ttk.Frame(win)
        btn_row.pack(pady=(4, 12))
        ttk.Button(btn_row, text="Clear", command=clear).pack(side="left", padx=5)
        ttk.Button(btn_row, text="Cancel", command=cancel).pack(side="left", padx=5)
        ttk.Button(btn_row, text="Accept", command=accept).pack(side="left", padx=5)

        win.wait_window()
        return result["path"]


class TopazSigner:
    """Captures a signature from a physical Topaz SigPlus pad via its COM
    control, if one is installed/connected on this machine. Raises
    TopazNotAvailable (never a raw exception) on any failure, so
    SignatureService can fall back to CanvasSigner cleanly."""

    def __init__(self, progid_options=None):
        self.progid_options = progid_options or ["TOPAZSIGPLUS.SigPlusCtrl.1", "SigPlus.SigPlusCtrl.1"]

    def _connect(self):
        import win32com.client as win32
        last_exc = None
        for progid in self.progid_options:
            try:
                sig = win32.Dispatch(progid)
                logger.info("Topaz: connected via ProgID %r", progid)
                return sig
            except Exception as exc:
                logger.debug("Topaz: ProgID %r failed: %s", progid, exc)
                last_exc = exc
        logger.error("Topaz: could not connect via any ProgID %r (last error: %s)",
                     self.progid_options, last_exc)
        raise TopazNotAvailable(
            f"Could not connect to a Topaz SigPlus pad ({last_exc}). "
            "Check the SigPlus SDK is installed and the pad is plugged in."
        )

    # ------------------------------------------------------------------
    # SigPlus builds do NOT all expose the same API, which is what broke
    # signing on the real office machine: the control connected fine via
    # ProgID 'SigPlus.SigPlusCtrl.1', then arming it blew up with
    #
    #   AttributeError: SigPlus.SigPlusCtrl.1.SetTabletState.
    #   Did you mean: 'SetTabletPortPath'?
    #
    # i.e. that build exposes TabletState as a PROPERTY (sig.TabletState
    # = 1), not as a SetTabletState(state, hwnd) METHOD. Older/other
    # builds do have the method. Rather than hard-coding either shape,
    # the helpers below try the known variants in turn and remember which
    # one worked, so the app adapts to whichever SigPlus is installed.
    # ------------------------------------------------------------------
    @staticmethod
    def _describe_api(sig):
        """Best-effort list of the method/property names the connected
        control actually exposes, for the log. Never raises."""
        try:
            type_info = sig._oleobj_.GetTypeInfo()
            attr = type_info.GetTypeAttr()
            names = set()
            for i in range(attr.cFuncs):
                try:
                    names.update(type_info.GetNames(type_info.GetFuncDesc(i).memid))
                except Exception:
                    pass
            for i in range(attr.cVars):
                try:
                    names.update(type_info.GetNames(type_info.GetVarDesc(i).memid))
                except Exception:
                    pass
            return sorted(n for n in names if n)
        except Exception as exc:
            return f"<could not enumerate control API: {exc}>"

    @staticmethod
    def _call_or_get(sig, name):
        """Reads a COM member that might be exposed as a plain property
        OR as a zero-arg method, depending on how this particular
        ActiveX build's type library declares it.

        Confirmed in the field: on the office laptop's SigPlus build,
        `sig.NumberOfTabletPoints` does not return a number - pywin32's
        dynamic dispatch hands back a bound method object instead
        (because that build's type library declares it as a callable
        member, not a property-get), so `points < 1` blew up with
        "'<' not supported between instances of 'method' and 'int'".
        Calling it (`sig.NumberOfTabletPoints()`) is what actually reads
        the value on that build. This helper does whichever is correct
        for whatever gets passed to it, so the same code works across
        SigPlus builds that expose a given member either way."""
        value = getattr(sig, name)
        if callable(value):
            value = value()
        return value

    @staticmethod
    def _set_tablet_state(sig, value, hwnd):
        """Arms (value=1) or disarms (value=0) the pad, whichever way
        this particular SigPlus build wants it. Returns the name of the
        variant that worked."""
        # Order matters: try the METHOD forms first. A missing method
        # fails loudly with AttributeError, whereas assigning a missing
        # property can silently succeed (it would just create a new
        # attribute), which would look like it worked while leaving the
        # pad untouched. So the property form is the last resort, and it
        # is read back afterwards to confirm it actually stuck.
        def _set_property():
            setattr(sig, "TabletState", value)
            try:
                readback = TopazSigner._call_or_get(sig, "TabletState")
            except Exception:
                return  # can't read it back; assume the put worked
            if readback != value:
                raise TopazNotAvailable(
                    f"TabletState did not stick (wrote {value!r}, read back {readback!r})"
                )

        attempts = (
            ("SetTabletState(state, hwnd)", lambda: sig.SetTabletState(value, hwnd)),
            ("SetTabletState(state)", lambda: sig.SetTabletState(value)),
            ("TabletState property", _set_property),
        )
        errors = []
        for label, call in attempts:
            try:
                call()
                logger.debug("Topaz: tablet state set to %s via %s", value, label)
                return label
            except Exception as exc:
                errors.append(f"{label}: {exc}")
        raise TopazNotAvailable(
            "This SigPlus build did not accept any known way of turning "
            "the pad on/off. Tried - " + " | ".join(errors)
        )

    @staticmethod
    def _read_signature_b64(sig):
        """Returns the captured signature as base64 PNG/BMP bytes if the
        control offers SigImageB64, else None (caller falls back to the
        file-based route)."""
        try:
            value = TopazSigner._call_or_get(sig, "SigImageB64")
            if value:
                logger.debug("Topaz: read signature via SigImageB64")
                return value
        except Exception as exc:
            logger.debug("Topaz: SigImageB64 not available (%s)", exc)
        return None

    @staticmethod
    def _write_signature_file(sig):
        """Fallback image route for builds without SigImageB64: ask the
        control to write a bitmap out, then hand back the path. Returns
        None if this build doesn't support it either.

        Office-laptop evidence (2026-09-08 log) settled the exact call
        shape this build needs:
            Topaz: image property ImageFileName not settable
                (Property 'SigPlus.SigPlusCtrl.1.ImageFileName' can not be set.)
            Topaz: WriteImageFile() failed
                ((-2147352561, 'Parameter not optional.', None, None))
            Topaz: WriteImageFile property failed
                (Property 'SigPlus.SigPlusCtrl.1.WriteImageFile' can not be set.)
        i.e. on THIS build ImageFileName is read-only (setting it silently
        no-ops on other builds, but here it flat out refuses), and
        WriteImageFile is a real method that REQUIRES its arguments to be
        passed directly - it does not read them back off the
        ImageFileName/ImageFileFormat/... properties the way older builds
        do. So we now try calling it with the path (and the other image
        settings) passed straight in as positional arguments, in a few
        decreasing arg-count shapes, before falling back to the old
        "set properties, call with no args" style for builds that DO
        want it that way."""
        out_path = os.path.join(tempfile.gettempdir(), f"sig_topaz_{uuid.uuid4().hex}.bmp")
        image_format, x_size, y_size, pen_width = 0, 500, 150, 2  # 0 = bitmap
        try:
            for prop, value in (
                ("ImageFileFormat", image_format),  # 0 = bitmap on the builds that have it
                ("ImageXSize", x_size),
                ("ImageYSize", y_size),
                ("ImagePenWidth", pen_width),
                ("JustifyMode", 5),          # NOT "ImageJustifyMode" - that name
                                              # doesn't exist on real SigPlus builds
                                              # (confirmed against the office
                                              # laptop's control API listing)
                ("ImageFileName", out_path),
            ):
                try:
                    setattr(sig, prop, value)
                except Exception as exc:
                    logger.debug("Topaz: image property %s not settable (%s)", prop, exc)
            wrote = False
            for label, call in (
                # This build: WriteImageFile() raised "Parameter not
                # optional" when called with zero args, so the filename
                # (and friends) must be passed directly as arguments -
                # try the fullest documented signature first, then
                # shorter ones, in case this build only wants the path.
                ("WriteImageFile(path, fmt, x, y, pen)",
                 lambda: sig.WriteImageFile(out_path, image_format, x_size, y_size, pen_width)),
                ("WriteImageFile(path, fmt)",
                 lambda: sig.WriteImageFile(out_path, image_format)),
                ("WriteImageFile(path)",
                 lambda: sig.WriteImageFile(out_path)),
                # Older-build style: properties already set above, call
                # with no args at all.
                ("WriteImageFile()", lambda: sig.WriteImageFile()),
                ("WriteImageFile property", lambda: setattr(sig, "WriteImageFile", 1)),
            ):
                try:
                    call()
                    wrote = True
                    logger.info("Topaz: wrote signature bitmap via %s", label)
                    break
                except Exception as exc:
                    # INFO (not DEBUG): this is the exact detail that
                    # settles WHICH call shape a given SigPlus build
                    # wants, and console-only screen recordings only show
                    # INFO+ - a DEBUG-only line here means a future field
                    # report can show the outer "no way to read the image
                    # back" error without any of the per-attempt detail
                    # that explains why, forcing a round trip just to ask
                    # for the raw log file.
                    logger.info("Topaz: %s failed (%s)", label, exc)
            if wrote and os.path.exists(out_path):
                return out_path
            if wrote:
                logger.debug(
                    "Topaz: WriteImageFile reported success but no file appeared at %s",
                    out_path,
                )
        except Exception:
            logger.exception("Topaz: file-based signature export failed")
        return TopazSigner._write_signature_via_bitmap_buffer(sig, out_path, x_size, y_size)

    @staticmethod
    def _write_signature_via_bitmap_buffer(sig, out_path, x_size, y_size):
        """Last-resort fallback for builds where WriteImageFile itself is
        unusable: read the captured signature straight out of the
        control's raw bitmap buffer (GetBitmapBufferBytes / ByteIndex /
        BitMapBufferByte / BitMapBufferClose - all confirmed present on
        the office laptop's control via check_topaz_api.py) and write it
        out as a BMP by hand. Returns the path, or None if this build
        doesn't expose that API either."""
        try:
            if not hasattr(sig, "GetBitmapBufferBytes"):
                return None
            n_bytes = TopazSigner._call_or_get(sig, "GetBitmapBufferBytes")
            n_bytes = int(n_bytes)
            if n_bytes <= 0:
                logger.debug("Topaz: GetBitmapBufferBytes returned %s, nothing to read", n_bytes)
                return None
            # Safety caps - this runs inside the same worker thread the
            # capture() timeout is watching, so don't let a huge/garbage
            # byte count (or a control that's just slow per-call over COM)
            # turn into a multi-minute read that hangs the app. A hard
            # byte-count ceiling AND a wall-clock budget both apply; either
            # one tripping aborts the read cleanly (falls back to on-screen
            # signing) rather than blocking indefinitely.
            MAX_BUFFER_BYTES = 400_000
            TIME_BUDGET_SECONDS = 8.0
            if n_bytes > MAX_BUFFER_BYTES:
                logger.debug(
                    "Topaz: GetBitmapBufferBytes=%s exceeds safety cap %s, skipping",
                    n_bytes, MAX_BUFFER_BYTES,
                )
                return None
            data = bytearray(n_bytes)
            start = time.time()
            for i in range(n_bytes):
                if time.time() - start > TIME_BUDGET_SECONDS:
                    logger.debug(
                        "Topaz: bitmap-buffer read exceeded %.1fs budget after %d/%d bytes, aborting",
                        TIME_BUDGET_SECONDS, i, n_bytes,
                    )
                    return None
                sig.ByteIndex = i
                data[i] = int(TopazSigner._call_or_get(sig, "BitMapBufferByte")) & 0xFF
            try:
                sig.BitMapBufferClose()
            except Exception as exc:
                logger.debug("Topaz: BitMapBufferClose() failed (%s)", exc)
            if data[:2] == b"BM":
                # The buffer is already a complete BMP file - write as-is.
                with open(out_path, "wb") as fh:
                    fh.write(bytes(data))
            else:
                # Raw pixel data with no file header - wrap it in a
                # minimal BMP header ourselves (matches ImageXSize/YSize
                # set on the control, 24-bit).
                from PIL import Image
                try:
                    img = Image.frombytes("RGB", (x_size, y_size), bytes(data))
                    img.save(out_path, "BMP")
                except Exception:
                    logger.exception(
                        "Topaz: could not interpret raw bitmap buffer as an image"
                    )
                    return None
            logger.info(
                "Topaz: wrote signature via raw bitmap buffer fallback (%d bytes)", n_bytes
            )
            return out_path if os.path.exists(out_path) else None
        except Exception:
            logger.exception("Topaz: bitmap-buffer fallback failed")
            return None

    def capture(self, parent, title="Sign on the pad"):
        import importlib.util
        if importlib.util.find_spec("win32com.client") is None:
            logger.error("Topaz: pywin32 (win32com.client) is not installed")
            raise TopazNotAvailable("pywin32 isn't installed.")
        import pythoncom
        logger.info("Topaz: capture() starting")

        # The SigPlus control is a classic STA ActiveX control: once it's
        # created on a thread, every later call into it must happen on
        # that SAME thread (COM apartment rules) - so the entire
        # connect -> arm -> wait -> read -> disarm sequence below runs
        # inside one dedicated worker thread, and the Tkinter dialog
        # stays on the main thread as always, the two only exchanging a
        # couple of small Events.
        #
        # This exists because a real Topaz pad was observed, in the
        # field, to make SetTabletState() block indefinitely - e.g. while
        # the device was still held open by another application (Acrobat
        # had just used it moments before). A direct, unguarded call froze
        # the whole app until the process was force-killed. Now the main
        # thread simply stops waiting after ARM_TIMEOUT_SECONDS and falls
        # back to on-screen signing instead; the stuck worker (a daemon
        # thread) is abandoned quietly rather than taking the app down.
        hwnd = parent.winfo_id()
        armed_event = threading.Event()
        proceed_event = threading.Event()
        done_event = threading.Event()
        state = {"sig": None, "connect_error": None}
        outcome = {"accepted": False, "b64": None, "path": None, "error": None}

        def worker():
            pythoncom.CoInitialize()
            try:
                try:
                    sig = self._connect()
                    logger.debug("Topaz: control exposes: %r", self._describe_api(sig))
                    logger.debug("Topaz: arming the pad (hwnd=%s)", hwnd)
                    self._set_tablet_state(sig, 1, hwnd)
                    try:
                        sig.ClearTablet()
                    except Exception as exc:
                        logger.debug("Topaz: ClearTablet() unavailable/failed (%s)", exc)
                    logger.info("Topaz: pad armed OK")
                except Exception as exc:
                    logger.exception("Topaz: failed to connect/arm the pad")
                    state["connect_error"] = exc
                    return
                state["sig"] = sig
            finally:
                armed_event.set()

            proceed_event.wait()
            sig = state["sig"]
            if sig is not None:
                if outcome["accepted"]:
                    try:
                        # NOT a plain attribute read - on the office
                        # laptop's SigPlus build, NumberOfTabletPoints is
                        # exposed as a zero-arg METHOD, not a property,
                        # and `sig.NumberOfTabletPoints` alone silently
                        # returns a bound-method object rather than a
                        # number. That made `points < 1` blow up with
                        # "'<' not supported between instances of
                        # 'method' and 'int'" and get miscounted as a
                        # captured-signature read failure. _call_or_get
                        # calls it if it's callable, reads it directly
                        # otherwise - works for both SigPlus shapes.
                        points = self._call_or_get(sig, "NumberOfTabletPoints")
                        logger.info("Topaz: NumberOfTabletPoints=%s after Accept", points)
                        if points < 1:
                            outcome["error"] = TopazNotAvailable("No signature was captured on the pad.")
                        else:
                            outcome["b64"] = self._read_signature_b64(sig)
                            if not outcome["b64"]:
                                # Build without SigImageB64 - have the
                                # control write a bitmap out instead.
                                outcome["path"] = self._write_signature_file(sig)
                                if not outcome["path"]:
                                    outcome["error"] = TopazNotAvailable(
                                        "The pad captured a signature, but this SigPlus build "
                                        "offered no way to read the image back (no SigImageB64 "
                                        "property and no working WriteImageFile)."
                                    )
                    except Exception as exc:
                        logger.exception("Topaz: could not read captured signature")
                        outcome["error"] = TopazNotAvailable(f"Could not read the captured signature ({exc}).")
                try:
                    self._set_tablet_state(sig, 0, hwnd)
                    logger.debug("Topaz: pad disarmed OK")
                except Exception:
                    logger.exception("Topaz: failed to disarm the pad (non-fatal)")
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            done_event.set()

        threading.Thread(target=worker, daemon=True).start()

        ARM_TIMEOUT_SECONDS = 6.0
        if not armed_event.wait(ARM_TIMEOUT_SECONDS):
            logger.error("Topaz: TIMED OUT after %.1fs waiting for pad to arm - "
                         "likely busy/held by another application", ARM_TIMEOUT_SECONDS)
            raise TopazNotAvailable(
                "Topaz pad did not respond within a few seconds - it may "
                "be busy or held open by another application (for "
                "example, still locked by Acrobat/Reader from a previous "
                "signing). Close any other program that might be using "
                "the pad and try again."
            )
        if state["connect_error"] is not None:
            raise TopazNotAvailable(
                f"Topaz pad did not respond ({state['connect_error']})."
            ) from state["connect_error"]

        win = tk.Toplevel(parent)
        win.title(title)
        win.grab_set()
        win.resizable(False, False)
        ttk.Label(win, text="Please sign on the Topaz pad now, then click Accept.",
                  wraplength=340, padding=20).pack()
        result = {"accepted": False}

        def accept():
            result["accepted"] = True
            win.destroy()

        def cancel():
            win.destroy()

        btn_row = ttk.Frame(win)
        btn_row.pack(pady=(0, 15))
        ttk.Button(btn_row, text="Cancel", command=cancel).pack(side="left", padx=5)
        ttk.Button(btn_row, text="Accept", command=accept).pack(side="left", padx=5)
        win.wait_window()

        outcome["accepted"] = result["accepted"]
        proceed_event.set()
        # Reading the captured points/image and disarming the pad is
        # normally near-instant - except the raw bitmap-buffer fallback
        # (used when neither SigImageB64 nor WriteImageFile work), which
        # can take up to its own ~8s time budget on builds that only
        # expose a slow, one-byte-at-a-time COM read. Give the whole
        # read+fallback+disarm sequence enough room for that worst case
        # plus slack, while still guarding against a truly unresponsive
        # pad freezing the app.
        done_event.wait(15.0)

        if not result["accepted"]:
            logger.info("Topaz: capture cancelled by user")
            return None
        if outcome["error"] is not None:
            logger.error("Topaz: capture failed: %s", outcome["error"])
            raise outcome["error"]
        if outcome["b64"] is None and outcome["path"] is None:
            logger.error("Topaz: no response from pad after Accept")
            raise TopazNotAvailable("Could not read the captured signature (no response from the pad).")

        try:
            if outcome["b64"] is not None:
                raw = base64.b64decode(outcome["b64"])
                img = Image.open(io.BytesIO(raw))
            else:
                # File-based route (build without SigImageB64): the
                # control wrote a bitmap, load it and convert to PNG so
                # the rest of the app sees the same thing either way.
                img = Image.open(outcome["path"])
                img.load()
        except Exception as exc:
            logger.exception("Topaz: could not decode captured signature image")
            raise TopazNotAvailable(f"Could not decode the captured signature image ({exc}).") from exc

        tmp_path = os.path.join(tempfile.gettempdir(), f"sig_topaz_{uuid.uuid4().hex}.png")
        img.save(tmp_path)
        logger.info("Topaz: capture SUCCESS, saved to %s", tmp_path)
        return tmp_path


class SignatureService:
    """Picks Topaz pad capture when CONFIG["use_topaz_pad"] is on and a
    pad actually responds, otherwise falls back to the on-screen
    CanvasSigner - so the wizard works identically on a machine with no
    pad attached (e.g. this one, for testing) and on a real walk-in kiosk
    with one connected."""

    def __init__(self, config_data):
        self.config_data = config_data
        self.canvas_signer = CanvasSigner()
        self.topaz_signer = TopazSigner(config_data.get("topaz_progid_options"))

    def capture(self, parent, title="Sign here"):
        use_topaz = self.config_data.get("use_topaz_pad", False)
        logger.info("SignatureService.capture: use_topaz_pad=%s, title=%r", use_topaz, title)
        if use_topaz:
            try:
                path = self.topaz_signer.capture(parent, title)
                logger.info("SignatureService.capture: Topaz path returned %r", path)
                return path
            except TopazNotAvailable as exc:
                logger.warning("Topaz pad not available, falling back to on-screen signing: %s", exc)
                messagebox.showwarning(
                    "Topaz pad not available",
                    f"{exc}\n\nFalling back to on-screen signature capture.",
                )
        path = self.canvas_signer.capture(parent, title)
        logger.info("SignatureService.capture: on-screen canvas path returned %r", path)
        return path


# ============================================================================
# CRYPTOGRAPHIC PDF SIGNING (pyHanko) - this is what replaces Adobe
# Acrobat + the Topaz plug-in's signing step. No Acrobat involved.
# ============================================================================

class SigningError(Exception):
    pass


def _signing_identity_paths(config_data):
    folder = _resolve_path(config_data, "signing_identity_folder", "signing_identity")
    return (
        os.path.join(folder, "signing_key.pem"),
        os.path.join(folder, "signing_cert.pem"),
    )


def ensure_signing_identity(config_data):
    """
    Creates a free, self-signed RSA/X.509 signing certificate + private key
    the FIRST time this app runs, and reuses the same one for every
    signature after that - this is what lets the app cryptographically
    sign PDFs by itself, with no Adobe Acrobat license, no Topaz plug-in,
    and no external Certificate Authority.

    TRADE-OFF (accepted): because nobody outside this organization vouches
    for this certificate, Acrobat/Reader will correctly show every PDF
    this app signs as genuinely signed and tamper-evident, but will also
    flag the certificate itself as "not trusted" - until/unless your IT
    team installs the cert this returns as a Trusted Certificate. See the
    README for exactly how.
    """
    key_path, cert_path = _signing_identity_paths(config_data)
    if os.path.exists(key_path) and os.path.exists(cert_path):
        return key_path, cert_path

    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "IT Asset Submission Acknowledgement"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Internal IT Department"),
    ])
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(datetime.utcnow() - timedelta(days=1))
        .not_valid_after(datetime.utcnow() + timedelta(days=3650))
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True, content_commitment=True, key_encipherment=False,
                data_encipherment=False, key_agreement=False, key_cert_sign=False,
                crl_sign=False, encipher_only=False, decipher_only=False,
            ),
            critical=True,
        )
        .sign(key, hashes.SHA256())
    )

    os.makedirs(os.path.dirname(key_path), exist_ok=True)
    with open(key_path, "wb") as f:
        f.write(key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        ))
    with open(cert_path, "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))
    return key_path, cert_path


def _find_signature_field_boxes(pdf_path, target_size_pt=(141.73, 47.24), tolerance=0.35):
    """
    Scans page 1 of `pdf_path` for placed images matching the fixed size
    every signature picture placeholder is always inserted at (see
    _picture_sdt() - 1800000x600000 EMU = 141.73x47.24pt, regardless of
    the BU template or the actual captured signature image's own pixel
    dimensions), and returns their on-page rects as (x0, y0, x1, y1) in
    PDF points, sorted top-to-bottom on the page.

    This lets sign_pdf_with_signatures() put a real, clickable signature
    FIELD widget exactly on top of the visible signature picture Word
    already rendered there - instead of an invisible field with no box,
    which is why a signature couldn't be clicked/inspected in Adobe
    Reader before this. Returns [] (never raises) if pypdf isn't
    installed or nothing on the page matches, so callers can fall back
    to an invisible field and signing still works either way.
    """
    try:
        from pypdf import PdfReader
        from pypdf.generic import ContentStream
    except ImportError:
        return []

    def mat_mult(m1, m2):
        # 2D affine matrices as PDF-style 6-tuples (a, b, c, d, e, f);
        # this composes m1 "then" m2, matching the PDF 'cm' operator's
        # convention of prepending the new matrix to the CTM.
        a1, b1, c1, d1, e1, f1 = m1
        a2, b2, c2, d2, e2, f2 = m2
        return (
            a1 * a2 + b1 * c2,
            a1 * b2 + b1 * d2,
            c1 * a2 + d1 * c2,
            c1 * b2 + d1 * d2,
            e1 * a2 + f1 * c2 + e2,
            e1 * b2 + f1 * d2 + f2,
        )

    def apply(m, x, y):
        a, b, c, d, e, f = m
        return (a * x + c * y + e, b * x + d * y + f)

    try:
        reader = PdfReader(pdf_path)
        page = reader.pages[0]
        content = ContentStream(page.get_contents(), reader)

        candidates = []
        ctm_stack = []
        ctm = (1, 0, 0, 1, 0, 0)
        for operands, operator in content.operations:
            op = operator.decode("latin1") if isinstance(operator, bytes) else operator
            if op == "q":
                ctm_stack.append(ctm)
            elif op == "Q":
                if ctm_stack:
                    ctm = ctm_stack.pop()
            elif op == "cm":
                m2 = tuple(float(v) for v in operands)
                ctm = mat_mult(m2, ctm)
            elif op == "Do":
                xobj_name = operands[0]
                try:
                    xobject = page["/Resources"]["/XObject"][xobj_name]
                    if xobject.get("/Subtype") == "/Image":
                        corners = [apply(ctm, x, y) for x, y in [(0, 0), (1, 0), (0, 1), (1, 1)]]
                        xs = [p[0] for p in corners]
                        ys = [p[1] for p in corners]
                        x0, x1 = min(xs), max(xs)
                        y0, y1 = min(ys), max(ys)
                        w, h = x1 - x0, y1 - y0
                        tw, th = target_size_pt
                        if abs(w - tw) <= tw * tolerance and abs(h - th) <= th * tolerance:
                            candidates.append((x0, y0, x1, y1))
                except Exception:
                    pass

        candidates.sort(key=lambda r: -r[3])  # top of page first
        return candidates
    except Exception:
        return []


def _stamp_signed_on_captions(pdf_path, boxes, when=None):
    """Enhancement 13/18 - draws a small 'Signed On: <date/time>' caption
    directly on the page, positioned just below each real signature box
    _find_signature_field_boxes() already detected. This is independent of
    the Word template (no template edit needed, and it doesn't touch/
    resize the actual signature picture the way stamping the caption onto
    the signature IMAGE itself would - that would have squeezed the visible
    signature to fit the template's fixed-size picture frame). Must run on
    the UNSIGNED pdf, same ordering rule as _set_pdf_metadata: editing page
    content after the cryptographic signature is applied would invalidate
    it. Never raises - a failure here should never block getting a signed
    PDF out the door; it just means this run doesn't get the caption."""
    if not HAS_REPORTLAB or not boxes:
        return
    try:
        from pypdf import PdfReader, PdfWriter

        reader = PdfReader(pdf_path)
        page = reader.pages[0]
        page_w = float(page.mediabox.width)
        page_h = float(page.mediabox.height)
        caption = f"Signed On: {(when or datetime.now()).strftime('%d-%b-%Y %I:%M %p')}"

        buf = io.BytesIO()
        c = _rl_canvas.Canvas(buf, pagesize=(page_w, page_h))
        c.setFont("Helvetica", 8)
        for (x0, y0, x1, _y1) in boxes:
            c.drawCentredString((x0 + x1) / 2.0, max(y0 - 11, 2), caption)
        c.save()
        buf.seek(0)

        page.merge_page(PdfReader(buf).pages[0])

        writer = PdfWriter()
        for p in reader.pages:
            writer.add_page(p)
        tmp_path = pdf_path + ".captiontmp"
        with open(tmp_path, "wb") as f:
            writer.write(f)
        os.replace(tmp_path, pdf_path)
        logger.info("Signed-on caption: stamped %d box(es) on %r", len(boxes), pdf_path)
    except Exception:
        logger.exception("Signed-on caption: could not stamp %r (non-fatal)", pdf_path)


def sign_pdf_with_signatures(input_pdf_path, output_pdf_path, signatures, config_data):
    """
    Applies one genuine, cryptographic PDF signature FIELD per entry in
    `signatures` (list of {"field_name", "display_name"}), each as its own
    incremental update to the file - the same underlying document
    structure Adobe Acrobat itself produces when signing with a Topaz pad,
    just produced by this app instead, using the self-signed certificate
    from ensure_signing_identity().

    Each field's on-page box is placed exactly on top of the VISIBLE
    signature picture the person actually drew/signed, which is already
    baked into the page by the Word->PDF export that ran before this - see
    fill_docx(). _find_signature_field_boxes() locates those pictures by
    their known fixed size; a blank StaticStampStyle (no border, no
    background/text) is used so nothing is drawn a second time on top of
    them - only a real, clickable signature field widget is added there.
    That's what makes the signature clickable/inspectable in Adobe
    Reader (Signature Properties, validity, etc.), instead of the field
    being invisible with nothing on the page to click. If the pictures
    can't be located for some reason (e.g. pypdf isn't installed, or a
    template was built without them), each field falls back to the old
    invisible (no on-page box) behavior so signing still succeeds.
    """
    logger.info(
        "Signing: starting for %r -> %r (%d signature field(s): %r)",
        input_pdf_path, output_pdf_path, len(signatures),
        [s.get("field_name") for s in signatures],
    )
    start_time = time.time()
    try:
        from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
        from pyhanko.sign import signers
        from pyhanko.sign.fields import SigFieldSpec
        from pyhanko.stamp import StaticStampStyle
    except ImportError as exc:
        logger.error("Signing: pyhanko/cryptography aren't installed")
        raise SigningError(
            "The 'pyhanko' library (and 'cryptography') aren't installed. Run:\n"
            "    pip install pyhanko cryptography"
        ) from exc

    key_path, cert_path = ensure_signing_identity(config_data)
    try:
        signer = signers.SimpleSigner.load(key_path, cert_path, key_passphrase=None)
    except Exception as exc:
        logger.exception("Signing: could not load the signing certificate")
        raise SigningError(f"Could not load this app's signing certificate: {exc}") from exc

    # Boxes are detected once, from the original unsigned PDF - the page's
    # content stream (and therefore the pictures' positions) doesn't
    # change across the incremental signing steps below, only new
    # objects get appended. Order matches `signatures` (Employee first,
    # then Asset Receiver), same as the top-to-bottom layout on the page.
    signature_boxes = _find_signature_field_boxes(input_pdf_path)
    logger.debug("Signing: detected %d on-page signature box(es)", len(signature_boxes))

    os.makedirs(os.path.dirname(output_pdf_path) or ".", exist_ok=True)
    current_input = input_pdf_path
    step_files = []
    try:
        for i, sig in enumerate(signatures):
            is_last = i == len(signatures) - 1
            step_output = output_pdf_path if is_last else f"{output_pdf_path}.step{i}.pdf"
            box = signature_boxes[i] if i < len(signature_boxes) else None
            with open(current_input, "rb") as inf:
                # strict=False: Microsoft Word's own PDF export writes
                # "hybrid" cross-reference sections (both a classic xref
                # table and an xref stream, for compatibility with older
                # readers) - pyHanko refuses those by default as a
                # precaution, but they're a completely normal, valid PDF
                # structure that Acrobat/Reader read from Word every day.
                writer = IncrementalPdfFileWriter(inf, strict=False)
                field_spec = SigFieldSpec(sig["field_name"], on_page=0, box=box)
                # Enhancement 17: reason/location, so Adobe's Signature
                # Properties panel shows real context instead of nothing -
                # this does NOT change whether the certificate itself
                # shows as "Trusted" (that still needs your IT team to
                # install it as a Trusted Certificate; see the README).
                sig_reason = sig.get("reason") or config_data.get("signing_reason") or None
                sig_location = sig.get("location") or None
                meta = signers.PdfSignatureMetadata(
                    field_name=sig["field_name"], name=sig.get("display_name") or None,
                    reason=sig_reason, location=sig_location,
                )
                # Blank appearance: box=None means an invisible field (no
                # appearance is drawn regardless), and a real box means
                # the visible picture Word already placed there is the
                # only thing that should show - no extra border/text on
                # top of it.
                stamp_style = StaticStampStyle(border_width=0, background=None) if box else None
                pdf_signer = signers.PdfSigner(
                    meta, signer, new_field_spec=field_spec, stamp_style=stamp_style
                )
                with open(step_output, "wb") as outf:
                    pdf_signer.sign_pdf(writer, output=outf)
            logger.debug(
                "Signing: applied field %r (box=%r) -> %r",
                sig["field_name"], box, step_output,
            )
            if not is_last:
                step_files.append(step_output)
            current_input = step_output
    except Exception as exc:
        logger.exception("Signing: FAILED after %.1fs", time.time() - start_time)
        raise SigningError(f"Could not apply the cryptographic signature: {exc}") from exc
    finally:
        for f in step_files:
            try:
                os.remove(f)
            except OSError:
                pass

    logger.info("Signing: SUCCESS -> %r (%.1fs)", output_pdf_path, time.time() - start_time)
    return output_pdf_path


# ============================================================================
# EMAIL DRAFT - pre-filled Outlook draft for LWD-style returns
# ============================================================================

def build_email_body(data):
    """Builds the plain-text body listing the employee and what was
    returned, from the same `data` dict used for the PDF."""
    assets = list(data.get("assets_issued", []))
    if data.get("assets_other"):
        assets.append(data["assets_other"])
    assets_line = ", ".join(assets) if assets else "(none selected)"

    lines = [
        "Hi,",
        "",
        "The following IT assets have been submitted/returned:",
        "",
        f"Employee ID: {data.get('emp_id', '')}",
        f"Employee Name: {data.get('emp_name', '')}",
        f"Manager Name: {data.get('manager_name', '')}",
        f"Submission Type: {data.get('submission_type', '')}",
        f"Action: {data.get('email_action_label') or 'Asset Return'}",
        f"Date: {data.get('date', '')}",
        "",
        f"Device Serial Number Returned: {data.get('current_device_serial', '')}",
        f"Assets Returned: {assets_line}",
        f"Asset Pending for Submission (if any): {data.get('asset_pending') or 'None'}",
        "",
        "Please find the signed acknowledgement form attached.",
        "",
        "Regards,",
    ]
    return "\n".join(lines)


def build_email_subject(data):
    # Enhancement 16: dedicated wording per submission type (CONFIG's
    # email_templates dict), while still defaulting to the exact original
    # text ("IT Asset Return - ...") when no action label was supplied.
    action_label = data.get("email_action_label") or "Return"
    return f"IT Asset {action_label} - {data.get('emp_name', '')} ({data.get('emp_id', '')})"


def open_draft(data, to_addresses="", attachment_path=None):
    """
    Opens a pre-filled Outlook draft window (NOT sent - the person reviews
    and clicks Send themselves). Raises RuntimeError with a clear message
    if Outlook isn't available so the caller can show it to the user.
    """
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "Could not open an Outlook draft: pywin32 isn't installed. "
            "Run 'pip install pywin32' (see requirements.txt)."
        ) from exc

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.Subject = build_email_subject(data)
        mail.Body = build_email_body(data)
        if to_addresses:
            mail.To = to_addresses
        if attachment_path:
            mail.Attachments.Add(attachment_path)
        mail.Display(False)  # opens the draft window - does NOT send it
    except Exception as exc:
        raise RuntimeError(
            f"Could not open an Outlook draft: {exc}\n"
            "Make sure Outlook is installed and has been opened/signed in "
            "at least once on this machine."
        ) from exc


# ============================================================================
# MAIN WIZARD APP
# ============================================================================

def _resolve_path(config_data, key, default):
    p = config_data.get(key, default)
    if not os.path.isabs(p):
        p = os.path.join(_app_dir(), p)
    return p


def _is_email_draft_type(config_data, submission_type):
    return submission_type in config_data.get("email_draft_submission_types", [])


def _safe_folder_name(name, fallback="Unspecified"):
    """Turns a submission type label ("Break Fix", "Contractor LWD", ...)
    into a filesystem-safe subfolder name, so generated PDFs land in
    generated_forms/<Submission Type>/ instead of all together in one
    flat folder."""
    name = (name or "").strip()
    if not name:
        return fallback
    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in name).strip()
    return safe.replace(" ", "_") or fallback


# ============================================================================
# PERSISTENT USER SETTINGS (Enhancements 2/3/4) - configured once (root
# save folder, operator name/location/BU/email), then reused automatically
# on every later launch instead of asking again. Stored per-Windows-user
# (their own profile folder), not next to app.py, so it survives an app.py
# update/replace and doesn't collide between different people sharing the
# same machine.
# ============================================================================

DEFAULT_USER_SETTINGS = {
    "root_save_folder": "",
    "operator_name": "",
    "location": "",
    "business_unit": "",
    "email": "",
}


def _user_settings_dir():
    """Where per-user settings live: %APPDATA%\\ITAssetSubmissionForm on
    Windows - the standard per-user, no-admin-required location.

    This is deliberately independent of _app_dir() (where logs/cache/
    signing_identity/bu_templates live, next to app.py or the .exe).
    That matters once this ships as a single-file .exe:
      - PyInstaller's --onefile build extracts itself to a throwaway
        temp folder (sys._MEIPASS) on every run and deletes it on exit -
        _app_dir() already correctly resolves to the real .exe's own
        folder instead (via sys.executable when sys.frozen is True), but
        settings still shouldn't live there.
      - The .exe itself gets replaced with a newer build periodically.
        %APPDATA% survives that untouched, since it isn't "next to the
        exe" at all.
      - The .exe may be deployed to a location normal users can't write
        to (e.g. Program Files) - %APPDATA% is always writable by the
        signed-in user, no admin rights needed.
      - Multiple people can share one machine/exe and each still gets
        their own settings, since %APPDATA% is already per-Windows-login.
    Falls back to the home directory on non-Windows (there is no APPDATA
    there)."""
    appdata = os.environ.get("APPDATA")
    base = appdata if appdata else os.path.expanduser("~")
    path = os.path.join(base, "ITAssetSubmissionForm")
    os.makedirs(path, exist_ok=True)
    return path


def _user_settings_path():
    return os.path.join(_user_settings_dir(), "user_settings.json")


# Where earlier builds of this app (before the %APPDATA% move above) used
# to store this file. _load_user_settings() migrates it automatically on
# first run after an update, so nobody who already ran the app loses
# their saved root folder/operator profile.
_LEGACY_USER_SETTINGS_PATH = os.path.join(os.path.expanduser("~"), ".it_asset_form_user_settings.json")


def _load_user_settings():
    path = _user_settings_path()
    if not os.path.exists(path) and os.path.exists(_LEGACY_USER_SETTINGS_PATH):
        try:
            with open(_LEGACY_USER_SETTINGS_PATH, encoding="utf-8") as f:
                legacy_data = json.load(f)
            logger.info(
                "User settings: migrating from legacy path %r -> %r",
                _LEGACY_USER_SETTINGS_PATH, path,
            )
            return _save_user_settings(legacy_data)
        except Exception:
            logger.exception("User settings: found a legacy settings file but could not migrate it, ignoring")
    if not os.path.exists(path):
        return dict(DEFAULT_USER_SETTINGS)
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        logger.exception("User settings: could not read %r, using defaults", path)
        return dict(DEFAULT_USER_SETTINGS)
    merged = dict(DEFAULT_USER_SETTINGS)
    merged.update({k: v for k, v in data.items() if k in DEFAULT_USER_SETTINGS})
    return merged


def _save_user_settings(data):
    path = _user_settings_path()
    to_write = dict(DEFAULT_USER_SETTINGS)
    to_write.update({k: v for k, v in data.items() if k in DEFAULT_USER_SETTINGS})
    with open(path, "w", encoding="utf-8") as f:
        json.dump(to_write, f, indent=2)
    logger.info(
        "User settings: saved to %r (root_save_folder=%r, operator_name=%r, business_unit=%r)",
        path, to_write.get("root_save_folder"), to_write.get("operator_name"), to_write.get("business_unit"),
    )
    return to_write


def _user_settings_configured(data):
    """True once the one piece of information every later launch actually
    depends on (root_save_folder) has been set. The rest (name/location/
    BU/email) are asked for on the same first-run screen but don't
    individually gate it - matching Enhancement 3's 'no repeated prompts'
    requirement."""
    return bool((data or {}).get("root_save_folder", "").strip())


# ============================================================================
# ADMIN / INSTALLATION SETTINGS - shared endpoints (SSRS report URL, AD
# Lookup URL, TracIT URL, ...) that apply to EVERYONE using this copy of the
# app, unlike user_settings.json above (which is per-Windows-login and lives
# in %APPDATA%). These live in admin_settings.json NEXT TO app.py/the .exe
# (via _app_dir(), same folder as bu_templates/, signing_identity/, logs/),
# so one edit there (or one edit in Settings by whoever manages the shared
# install) applies to every person who runs that copy.
#
# IMPORTANT: the AD-lookup helper is launched as a completely separate
# subprocess (--webview-helper, see _run_webview_helper()) that re-imports
# this module fresh and reads the module-level CONFIG dict directly - it has
# no access to a running WizardApp's self.config_data. So these overrides
# are applied at the MODULE level (mutating CONFIG in place) via
# _apply_admin_settings_overrides(), and that function is called from BOTH
# WizardApp.__init__ (main GUI process) and _run_webview_helper() (helper
# subprocess) independently - see each call site for why.
# ============================================================================

ADMIN_SETTINGS_FIELDS = {
    # admin_settings.json key -> (CONFIG dict path, human label)
    "ssrs_report_url": (("ssrs_asset_report", "url"), "SSRS Asset Report URL"),
    "ad_lookup_url_template": (("ad_lookup_url_template",), "AD Lookup URL template"),
    "tracit_url_template": (("tracit_url_template",), "TracIT URL template"),
    "ssrs_cache_max_age_hours": (("ssrs_asset_report", "cache_max_age_hours"), "SSRS cache max age (hours)"),
}


def _admin_settings_path():
    return os.path.join(_app_dir(), "admin_settings.json")


def _load_admin_settings():
    """Reads admin_settings.json next to the exe. Missing file or missing
    individual keys are both fine - _apply_admin_settings_overrides() only
    overrides what's actually present, so CONFIG's built-in defaults keep
    working until someone deliberately sets something in Settings."""
    path = _admin_settings_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return {k: v for k, v in (data or {}).items() if k in ADMIN_SETTINGS_FIELDS}
    except Exception:
        logger.exception("Admin settings: could not read %r, ignoring overrides", path)
        return {}


def _save_admin_settings(data):
    path = _admin_settings_path()
    to_write = {k: v for k, v in (data or {}).items() if k in ADMIN_SETTINGS_FIELDS and str(v).strip() != ""}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(to_write, f, indent=2)
    logger.info("Admin settings: saved to %r (%s)", path, ", ".join(sorted(to_write.keys())) or "no overrides set")
    return to_write


def _apply_admin_settings_overrides(config, admin_data=None):
    """Mutates `config` in place, overriding the built-in default endpoints
    with anything set in admin_settings.json. Safe to call with an empty/
    missing admin_data (a no-op) so every startup path can call it
    unconditionally."""
    admin_data = admin_data if admin_data is not None else _load_admin_settings()
    for key, value in (admin_data or {}).items():
        if value is None or str(value).strip() == "":
            continue
        mapping = ADMIN_SETTINGS_FIELDS.get(key)
        if not mapping:
            continue
        path, _label = mapping
        node = config
        for part in path[:-1]:
            node = node.setdefault(part, {})
        leaf = path[-1]
        if leaf == "cache_max_age_hours":
            try:
                value = float(value)
            except (TypeError, ValueError):
                continue
        node[leaf] = value
    return config


# ============================================================================
# AUDIT LOG (Enhancement 14) - one line per PDF-generation attempt
# (success or failure), separate from the diagnostic app.log above and
# never rotated/truncated automatically. Plain CSV so it opens directly
# in Excel for a quick audit review.
# ============================================================================

AUDIT_LOG_COLUMNS = [
    "timestamp", "operator", "employee_id", "submission_type",
    "ssrs_match_status", "lookup_source", "pdf_generated", "pdf_location", "email_draft_created",
]


def _audit_log_path():
    log_dir = os.path.join(_app_dir(), "logs")
    os.makedirs(log_dir, exist_ok=True)
    return os.path.join(log_dir, "audit.log")


def _write_audit_log_entry(operator="", employee_id="", submission_type="",
                            ssrs_match_status="", lookup_source="", pdf_generated=False,
                            pdf_location="", email_draft_created=False):
    path = _audit_log_path()
    is_new = not os.path.exists(path)
    try:
        with open(path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if is_new:
                writer.writerow(AUDIT_LOG_COLUMNS)
            writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                operator, employee_id, submission_type, ssrs_match_status, lookup_source,
                "Yes" if pdf_generated else "No", pdf_location,
                "Yes" if email_draft_created else "No",
            ])
    except Exception:
        logger.exception("Audit log: could not write entry (employee_id=%r)", employee_id)


def _read_audit_log_entries():
    """Returns a list of dicts (one per row), newest first, for the
    Generated Documents history screen. Never raises - returns [] on any
    problem reading the file (e.g. it doesn't exist yet)."""
    path = _audit_log_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
        rows.reverse()
        return rows
    except Exception:
        logger.exception("Audit log: could not read %r", path)
        return []


# ============================================================================
# PDF METADATA (Enhancement 19) - embedded via pypdf, on the UNSIGNED pdf
# only (see the call site in _on_generate/_batch_save_and_complete) -
# editing bytes after the cryptographic signature has been applied would
# invalidate it.
# ============================================================================

def _set_pdf_metadata(pdf_path, meta):
    """Writes standard PDF /Info metadata into pdf_path IN PLACE. Any
    error here is logged, never raised - metadata is a nice-to-have, not
    something that should ever block getting a signed PDF out the door."""
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        logger.warning("PDF metadata: pypdf not installed, skipping")
        return
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        writer.append(reader)
        writer.add_metadata({f"/{k}": str(v) for k, v in meta.items() if v})
        tmp_path = pdf_path + ".metatmp"
        with open(tmp_path, "wb") as f:
            writer.write(f)
        os.replace(tmp_path, pdf_path)
        logger.debug("PDF metadata: wrote %d field(s) into %r", len(meta), pdf_path)
    except Exception:
        logger.exception("PDF metadata: failed to write into %r (non-fatal)", pdf_path)


# ============================================================================
# SSRS ASSET REPORT (Enhancements 7-12) - downloaded once per day (or on
# demand via "Refresh Data"), cached locally, and indexed by Employee ID
# so New Hire asset details can be auto-filled instead of typed in by
# hand. Uses the current Windows login session (SSPI) - no stored
# password, same mechanism as the existing "direct" AD Lookup mode above.
# ============================================================================

SSRS_STATUS_NOT_AVAILABLE = "Not Available"
SSRS_STATUS_AVAILABLE = "Available"
SSRS_STATUS_DOWNLOAD_FAILED = "Download Failed"
SSRS_STATUS_STALE = "Stale"


def _ssrs_cache_path():
    cache_dir = os.path.join(_app_dir(), "cache")
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, "ssrs_asset_report.xlsx")


def _ssrs_cache_is_fresh(path, max_age_hours=24):
    """Matches Enhancement 9's 'Downloaded Today?' check: fresh only if
    the cached file exists, was last modified today, AND within
    max_age_hours - so a file from just before midnight still counts as
    stale first thing the next morning even though max_age_hours alone
    wouldn't have expired it yet."""
    if not os.path.exists(path):
        return False
    mtime = datetime.fromtimestamp(os.path.getmtime(path))
    if mtime.date() != datetime.now().date():
        return False
    age_hours = (datetime.now() - mtime).total_seconds() / 3600.0
    return age_hours <= max_age_hours


def download_ssrs_asset_report(url, dest_path, timeout=60):
    """Downloads the SSRS report (an .xlsx, via rs:Format=EXCELOPENXML) to
    dest_path using the current Windows login session - no credentials
    are ever entered or stored by this app. Raises RuntimeError with a
    clear, human-readable message on any failure."""
    if requests is None:
        raise RuntimeError(
            "The 'requests' library is not installed. Run:\n    pip install requests"
        )
    if not HAS_SSPI:
        raise RuntimeError(
            "The 'requests_negotiate_sspi' library is not installed, so this app "
            "can't authenticate to SSRS with your current Windows login. Run:\n"
            "    pip install requests-negotiate-sspi\n"
            "(Windows only - this lets the download reuse your existing network "
            "login instead of a password prompt.)"
        )
    auth = HttpNegotiateAuth()
    try:
        resp = requests.get(url, auth=auth, timeout=timeout)
        resp.raise_for_status()
    except requests.exceptions.SSLError as exc:
        # NOTE: must be caught BEFORE requests.exceptions.ConnectionError below,
        # since SSLError is a subclass of it - otherwise this more specific,
        # more actionable message would never be reached.
        truststore_hint = (
            "" if HAS_TRUSTSTORE else
            "\n\nThis app tried to install the 'truststore' package fix for this "
            "automatically but it isn't installed. Ask IT/your admin to run:\n"
            "    pip install truststore\n"
            "then restart the app."
        )
        raise RuntimeError(
            "SSRS download failed because of a certificate trust problem, not a "
            "network/VPN problem:\n\n"
            f"{exc}\n\n"
            "This usually means your company's internal root certificate is "
            "trusted by Windows/your browser (which is why the report downloads "
            "fine in Chrome) but is not in Python's separate list of trusted "
            f"certificates.{truststore_hint}"
        ) from exc
    except requests.exceptions.ConnectionError as exc:
        raise RuntimeError(
            f"Could not reach the SSRS report server:\n{url}\n\n"
            "Check that you're connected to the corporate network/VPN."
        ) from exc
    except requests.exceptions.Timeout as exc:
        raise RuntimeError("The SSRS report server took too long to respond.") from exc
    except requests.exceptions.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "?"
        raise RuntimeError(
            f"SSRS report server returned an error (HTTP {status}). You may not "
            "have permission to run this report, or the URL may have changed - "
            "check with the report owner."
        ) from exc
    except Exception as exc:
        raise RuntimeError(f"SSRS download failed: {exc}") from exc

    os.makedirs(os.path.dirname(dest_path) or ".", exist_ok=True)
    tmp_path = dest_path + ".part"
    with open(tmp_path, "wb") as f:
        f.write(resp.content)
    os.replace(tmp_path, dest_path)
    return dest_path


def _find_ssrs_header_row(rows, required_columns, max_scan_rows=6):
    """Same lesson learned from the AD lookup 'sizer row' bug: don't
    assume row 0 of an SSRS export is always the real header row - scan
    the first few rows for the one that actually contains at least one of
    the expected column names. Returns the 0-based row index, or None."""
    wanted = {str(c).strip().lower() for c in required_columns}
    for i, row in enumerate(rows[:max_scan_rows]):
        cells = {str(c).strip().lower() for c in row if c is not None}
        if wanted & cells:
            return i
    return None


def parse_ssrs_asset_workbook(xlsx_path, column_map):
    """Reads xlsx_path (an SSRS EXCELOPENXML export) and returns
    (index, record_count) where index is {employee_id_string: {app_field:
    value, ...}}, built from column_map ({app_field: "SSRS Column Name"}).
    Each row is indexed under both its literal Employee ID text AND a
    digits-only variant, so a lookup by either '0012345' or '12345' finds
    the same row - the same lesson learned from the AD Lookup Employee ID
    field."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        header_scan_rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=6, values_only=True)]

        header_row_idx = _find_ssrs_header_row(header_scan_rows, column_map.values())
        if header_row_idx is None:
            raise RuntimeError(
                "Downloaded the SSRS report, but couldn't find the expected column "
                f"headers ({', '.join(column_map.values())}) in its first few rows. "
                "The report layout may have changed - check ssrs_asset_report."
                "column_map in CONFIG against the real file."
            )

        header_cells = [str(c).strip() if c is not None else "" for c in header_scan_rows[header_row_idx]]
        col_index = {}
        for app_field, ssrs_col in column_map.items():
            try:
                col_index[app_field] = header_cells.index(ssrs_col)
            except ValueError:
                continue  # column not present in this export - field left unpopulated

        if "employee_id" not in col_index:
            raise RuntimeError(
                f"Downloaded the SSRS report, but its Employee ID column "
                f"({column_map.get('employee_id')!r}) wasn't found among the "
                f"headers: {header_cells}"
            )

        index = {}
        count = 0
        for row in ws.iter_rows(min_row=header_row_idx + 2, values_only=True):
            if row is None or all(v is None or str(v).strip() == "" for v in row):
                continue
            emp_idx = col_index["employee_id"]
            emp_id_cell = row[emp_idx] if emp_idx < len(row) else None
            emp_id = str(emp_id_cell).strip() if emp_id_cell is not None else ""
            if not emp_id:
                continue
            record = {}
            for app_field, idx in col_index.items():
                if app_field == "employee_id" or idx >= len(row):
                    continue
                val = row[idx]
                record[app_field] = "" if val is None else str(val).strip()
            index[emp_id] = record
            digits_only = "".join(c for c in emp_id if c.isdigit())
            if digits_only and digits_only != emp_id:
                index.setdefault(digits_only, record)
            count += 1
        return index, count
    finally:
        wb.close()


# ============================================================================
# STARTUP HEALTH CHECK (Enhancement 13)
# ============================================================================

def _check_word_installed():
    """Best-effort, non-invasive check - reads the registry, does NOT
    launch Word. True if Microsoft Word's COM ProgID is registered.
    Always returns False (never raises) on non-Windows or on any error -
    this is a startup diagnostic, not something that should crash the
    app."""
    if sys.platform != "win32":
        return False
    try:
        import winreg
        winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, "Word.Application")
        return True
    except Exception:
        return False


def run_startup_health_check(config_data, user_settings, ssrs_status):
    """Returns an ordered list of (label, ok, detail) tuples for the
    startup health-check panel. Every check is wrapped in its own
    try/except so one failing check can never prevent the others (or the
    app itself) from proceeding."""
    checks = []

    try:
        word_ok = _check_word_installed()
        checks.append((
            "Microsoft Word", word_ok,
            "" if word_ok else "Word.Application COM ProgID not found in the registry.",
        ))
    except Exception as exc:
        checks.append(("Microsoft Word", False, str(exc)))

    try:
        _, cert_path = _signing_identity_paths(config_data)
        ok = os.path.exists(cert_path)
        checks.append((
            "Signing Certificate", ok,
            "" if ok else f"Not created yet at {cert_path} (generated automatically on first use).",
        ))
    except Exception as exc:
        checks.append(("Signing Certificate", False, str(exc)))

    checks.append(("SSRS Asset Report", ssrs_status == SSRS_STATUS_AVAILABLE, ssrs_status))

    try:
        root = (user_settings or {}).get("root_save_folder") or _resolve_path(
            config_data, "pdf_save_folder", "generated_forms"
        )
        os.makedirs(root, exist_ok=True)
        ok = os.access(root, os.W_OK)
        checks.append(("Save Folder", ok, root if ok else f"{root} is not writable."))
    except Exception as exc:
        checks.append(("Save Folder", False, str(exc)))

    try:
        templates = list_templates(_resolve_path(config_data, "bu_templates_folder", "bu_templates"))
        ok = len(templates) > 0
        checks.append((
            "BU Templates", ok,
            f"{len(templates)} loaded" if ok else "None configured yet - add one via 'Add New BU Template...'.",
        ))
    except Exception as exc:
        checks.append(("BU Templates", False, str(exc)))

    return checks


class WizardApp(tk.Tk):
    """Two-tab app: 'Single Person' (one-off form, everything on one
    scrollable page) and 'Bulk Batch' (import an Excel list, see everyone
    in a live-status list, and process them in any order - or let the
    app auto-advance through them)."""

    def __init__(self):
        super().__init__()
        self.config_data = copy.deepcopy(CONFIG)  # settings live at the top of this file
        self.admin_settings = _load_admin_settings()
        _apply_admin_settings_overrides(self.config_data, self.admin_settings)
        self.user_settings = _load_user_settings()
        self.ssrs_state = {
            "status": SSRS_STATUS_NOT_AVAILABLE, "records": 0,
            "last_updated": None, "index": {}, "error": "",
        }
        logger.info(
            "WizardApp: starting up. ad_lookup_mode=%r use_topaz_pad=%r bu_templates_folder=%r "
            "root_save_folder=%r operator_name=%r",
            self.config_data.get("ad_lookup_mode"),
            self.config_data.get("use_topaz_pad"),
            self.config_data.get("bu_templates_folder"),
            self.user_settings.get("root_save_folder"),
            self.user_settings.get("operator_name"),
        )
        self.browser_session = BrowserLookupSession()
        self.signature_service = SignatureService(self.config_data)

        self.title("IT Asset Submission Acknowledgement")
        # Size to the ACTUAL screen instead of a fixed "950x860" - on a
        # laptop with a smaller or scaled display, a hardcoded height can
        # end up taller than the visible screen, pushing the bottom
        # Save/Generate button bar off-screen with no way to reach it
        # (see _set_windows_dpi_awareness for the other half of this fix).
        # Opening maximized by default is the most robust guarantee that
        # the bottom bar is always visible, whatever the screen size.
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        want_w, want_h = 950, 860
        fit_w = min(want_w, screen_w - 40)
        fit_h = min(want_h, screen_h - 80)  # leave room for the taskbar
        self.geometry(f"{fit_w}x{fit_h}")
        self.minsize(min(860, fit_w), min(620, fit_h))
        try:
            self.state("zoomed")  # Windows: start maximized
        except Exception:
            logger.debug("WizardApp: could not start maximized (non-fatal)", exc_info=True)

        self._bu_templates_cache = []

        self._apply_ui_theme()
        self._build_shell()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_single_tab_shell()
        self._build_form()

        self._build_bulk_tab_shell()
        self._build_history_tab_shell()

        if _uses_webview_lookup(self.config_data):
            threading.Thread(target=self.browser_session.start, daemon=True).start()

        # First-run setup (Enhancements 2-4) happens after the shell/tabs
        # exist (so the dialog has a real parent window and _refresh_save_path
        # is already wired up) but before the person starts using the form.
        if not _user_settings_configured(self.user_settings):
            self.after(200, lambda: self._show_settings_dialog(first_run=True))

        # Kicks off a background thread; the SSRS status panel and startup
        # health check update live as soon as it finishes - neither blocks
        # the app from being usable in the meantime.
        self._ssrs_refresh(force=False)
        self.after(400, self._show_startup_health_check)

        logger.info("WizardApp: GUI ready")

    # ------------------------------------------------------------ look & feel
    def _apply_ui_theme(self):
        """A bounded visual-polish pass: consistent fonts, an Optum-orange
        accent color on primary actions/headers, and light section framing.
        Deliberately NOT a layout rewrite - every existing widget, grid/pack
        call and callback is untouched; this only changes ttk.Style
        defaults, so nothing about how the form behaves changes."""
        ACCENT = "#EB690B"       # Optum orange
        ACCENT_DARK = "#C4550A"
        BG = "#F7F7F5"
        SURFACE = "#FFFFFF"
        TEXT = "#222222"
        MUTED = "#666666"
        BORDER = "#DDDDDD"

        try:
            self.configure(background=BG)
        except Exception:
            pass

        style = ttk.Style(self)
        try:
            # 'clam' renders custom colors far more reliably on Windows than
            # the default 'vista'/'winnative' theme, which ignores most
            # ttk.Style color overrides.
            style.theme_use("clam")
        except Exception:
            logger.debug("UI theme: 'clam' unavailable, keeping default theme", exc_info=True)

        base_font = ("Segoe UI", 10)
        try:
            self.option_add("*Font", base_font)
        except Exception:
            pass

        style.configure(".", background=BG, foreground=TEXT, font=base_font)
        style.configure("TFrame", background=BG)
        style.configure("TLabelframe", background=BG, bordercolor=BORDER)
        style.configure("TLabelframe.Label", background=BG, foreground=ACCENT_DARK, font=("Segoe UI", 10, "bold"))
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("TCheckbutton", background=BG)
        style.configure("TRadiobutton", background=BG)
        style.configure("TSeparator", background=BORDER)

        style.configure("TEntry", fieldbackground=SURFACE, bordercolor=BORDER)
        style.configure("TCombobox", fieldbackground=SURFACE)

        style.configure(
            "TButton", background=ACCENT, foreground="white",
            font=("Segoe UI", 10, "bold"), padding=(10, 5), borderwidth=0,
        )
        style.map(
            "TButton",
            background=[("active", ACCENT_DARK), ("disabled", "#C9C9C9")],
            foreground=[("disabled", "#888888")],
        )

        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure(
            "TNotebook.Tab", background="#EDEDEA", foreground=TEXT,
            font=("Segoe UI", 10, "bold"), padding=(16, 8),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", SURFACE)],
            foreground=[("selected", ACCENT_DARK)],
        )

        # Status bar at the bottom - kept visually distinct/muted rather
        # than looking like an editable field.
        style.configure("Status.TLabel", background="#EDEDEA", foreground=MUTED, padding=(6, 3))

    # ------------------------------------------------------------ tooltips
    def _add_info_tooltip(self, parent, text, side="left", padx=(4, 0)):
        """A small 'ⓘ' label that shows `text` in a borderless popup on
        hover - lets a section header stay a couple of words (e.g.
        'Business Unit') while the longer explanation most people only
        need once is still one hover away instead of always taking up
        vertical space on screen."""
        icon = ttk.Label(parent, text="ⓘ", foreground="#3277c9", cursor="hand2")
        icon.pack(side=side, padx=padx)
        tip = {"win": None}

        def show(_event=None):
            if tip["win"] is not None:
                return
            win = tk.Toplevel(self)
            win.wm_overrideredirect(True)
            win.wm_attributes("-topmost", True)
            x = icon.winfo_rootx() + 4
            y = icon.winfo_rooty() + icon.winfo_height() + 4
            win.wm_geometry(f"+{x}+{y}")
            ttk.Label(
                win, text=text, wraplength=340, justify="left", padding=8,
                background="#FFFDE7", relief="solid", borderwidth=1,
            ).pack()
            tip["win"] = win

        def hide(_event=None):
            if tip["win"] is not None:
                tip["win"].destroy()
                tip["win"] = None

        icon.bind("<Enter>", show)
        icon.bind("<Leave>", hide)
        return icon

    # -------------------------------------------------------- keyboard nav
    @staticmethod
    def _bind_enter_advances_focus(widget):
        """Enhancement 8 - Tab/Shift+Tab already follow the natural widget
        creation order in Tk with no extra code needed; this adds Enter as
        an equivalent 'move to the next field' for a plain text Entry that
        doesn't already have its own Enter behavior (Employee ID submits a
        lookup on Enter, for example - that's intentionally left alone)."""
        widget.bind("<Return>", lambda e: (e.widget.tk_focusNext().focus(), "break")[-1])

    # ------------------------------------------------------------ shell
    def _build_shell(self):
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, anchor="w", style="Status.TLabel").pack(fill="x", side="bottom")

        self._build_header_bar()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.single_tab = ttk.Frame(self.notebook)
        self.bulk_tab = ttk.Frame(self.notebook)
        self.history_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.single_tab, text="Single Person")
        self.notebook.add(self.bulk_tab, text="Bulk Batch")
        self.notebook.add(self.history_tab, text="Generated Documents")

    # ------------------------------------------------------- branding/header
    def _build_header_bar(self):
        """Enhancement 1 (Optum branding) + the SSRS Asset Report status
        dashboard (Enhancements 10/11) + a 'Settings...' button to reopen
        the one-time setup dialog (Enhancement 2-4) any time."""
        header = ttk.Frame(self, padding=(10, 6))
        header.pack(fill="x", side="top")

        brand = ttk.Frame(header)
        brand.pack(side="left")
        self._header_logo_image = None  # kept as an attribute so Tk doesn't GC it
        logo_path = os.path.join(_app_dir(), "assets", "optum_logo.png")
        if os.path.exists(logo_path):
            try:
                img = Image.open(logo_path)
                img.thumbnail((140, 40))
                self._header_logo_image = ImageTk.PhotoImage(img)
                ttk.Label(brand, image=self._header_logo_image).pack(side="left", padx=(0, 10))
            except Exception:
                logger.exception("Header: could not load logo at %r (falling back to text)", logo_path)
        if self._header_logo_image is None:
            # No real logo file provided yet - drop assets/optum_logo.png next
            # to app.py to show the actual logo here instead of this text
            # fallback. PDF templates already contain the real logo and are
            # untouched either way.
            ttk.Label(brand, text="OPTUM", font=("Segoe UI", 16, "bold"), foreground="#EB690B").pack(
                side="left", padx=(0, 10)
            )
        title_box = ttk.Frame(brand)
        title_box.pack(side="left")
        ttk.Label(
            title_box, text="IT Asset Submission Acknowledgement System", font=("Segoe UI", 11, "bold")
        ).pack(anchor="w")
        self.header_operator_var = tk.StringVar(value=self._format_operator_line())
        ttk.Label(title_box, textvariable=self.header_operator_var, foreground="#666").pack(anchor="w")

        right = ttk.Frame(header)
        right.pack(side="right")

        ssrs_box = ttk.LabelFrame(right, text="SSRS Asset Report", padding=6)
        ssrs_box.pack(side="left", padx=(0, 10))
        self.ssrs_status_var = tk.StringVar(value=f"Status: {SSRS_STATUS_NOT_AVAILABLE}")
        self.ssrs_records_var = tk.StringVar(value="Records Loaded: 0")
        self.ssrs_updated_var = tk.StringVar(value="Last Updated: never")
        ttk.Label(ssrs_box, textvariable=self.ssrs_status_var).pack(anchor="w")
        ttk.Label(ssrs_box, textvariable=self.ssrs_records_var).pack(anchor="w")
        ttk.Label(ssrs_box, textvariable=self.ssrs_updated_var).pack(anchor="w")
        ttk.Button(ssrs_box, text="Refresh Data", command=lambda: self._ssrs_refresh(force=True)).pack(
            anchor="w", pady=(4, 0)
        )

        ttk.Button(right, text="Settings...", command=lambda: self._show_settings_dialog(first_run=False)).pack(
            side="left"
        )

        ttk.Separator(self, orient="horizontal").pack(fill="x", side="top")

    def _format_operator_line(self):
        name = self.user_settings.get("operator_name") or ""
        location = self.user_settings.get("location") or ""
        if not name and not location:
            return "Operator not set up yet - click Settings to configure."
        return "Operator: " + name + (f"  |  {location}" if location else "")

    def _refresh_header_operator_label(self):
        if hasattr(self, "header_operator_var"):
            self.header_operator_var.set(self._format_operator_line())

    def _refresh_ssrs_status_labels(self):
        if not hasattr(self, "ssrs_status_var"):
            return
        state = self.ssrs_state
        self.ssrs_status_var.set(f"Status: {state.get('status', SSRS_STATUS_NOT_AVAILABLE)}")
        self.ssrs_records_var.set(f"Records Loaded: {state.get('records', 0)}")
        last_updated = state.get("last_updated")
        self.ssrs_updated_var.set(
            "Last Updated: " + (last_updated.strftime("%d-%b-%Y %I:%M %p") if last_updated else "never")
        )
        self.status_var.set(
            f"SSRS asset report: {state.get('status')}"
            + (f" - {state['error']}" if state.get("error") else "")
        )

    def _ssrs_refresh(self, force=False):
        """Downloads/re-parses the SSRS asset report in a background
        thread (Enhancements 9/11) so the UI never blocks on the network
        call, then hops back to the main thread (self.after) to update
        state and the status labels - Tkinter widgets must only be
        touched from the main thread."""

        def worker():
            cfg = self.config_data.get("ssrs_asset_report", {})
            url = cfg.get("url", "")
            column_map = cfg.get("column_map", {})
            max_age = cfg.get("cache_max_age_hours", 24)
            cache_path = _ssrs_cache_path()
            error_detail = ""

            if force or not _ssrs_cache_is_fresh(cache_path, max_age):
                try:
                    if not url:
                        raise RuntimeError("ssrs_asset_report.url is not configured.")
                    logger.info("SSRS asset report: downloading (force=%r) -> %r", force, cache_path)
                    download_ssrs_asset_report(url, cache_path)
                except Exception as exc:
                    logger.exception("SSRS asset report: download failed (will try existing cache, if any)")
                    error_detail = str(exc)

            index, records, status, last_updated = {}, 0, SSRS_STATUS_NOT_AVAILABLE, None
            if os.path.exists(cache_path):
                try:
                    index, records = parse_ssrs_asset_workbook(cache_path, column_map)
                    fresh = _ssrs_cache_is_fresh(cache_path, max_age)
                    status = SSRS_STATUS_AVAILABLE if (fresh and not error_detail) else SSRS_STATUS_STALE
                    last_updated = datetime.fromtimestamp(os.path.getmtime(cache_path))
                except Exception as exc:
                    logger.exception("SSRS asset report: could not parse cached file %r", cache_path)
                    error_detail = error_detail or str(exc)
                    status = SSRS_STATUS_DOWNLOAD_FAILED
            else:
                status = SSRS_STATUS_DOWNLOAD_FAILED if error_detail else SSRS_STATUS_NOT_AVAILABLE

            def apply():
                self.ssrs_state = {
                    "status": status, "records": records,
                    "last_updated": last_updated, "index": index, "error": error_detail,
                }
                self._refresh_ssrs_status_labels()
                logger.info(
                    "SSRS asset report: status=%r records=%d error=%r",
                    status, records, error_detail,
                )

            self.after(0, apply)

        threading.Thread(target=worker, daemon=True).start()

    def _show_startup_health_check(self):
        """Enhancement 13 - a quick, non-blocking startup readiness panel.
        The SSRS line reflects whatever the local cache looks like right
        now (a fast, synchronous, no-network check) rather than waiting
        on the background download kicked off above - the SSRS Asset
        Report box in the header is the live/authoritative status."""
        cache_path = _ssrs_cache_path()
        max_age = self.config_data.get("ssrs_asset_report", {}).get("cache_max_age_hours", 24)
        if not os.path.exists(cache_path):
            ssrs_snapshot = SSRS_STATUS_NOT_AVAILABLE
        elif _ssrs_cache_is_fresh(cache_path, max_age):
            ssrs_snapshot = SSRS_STATUS_AVAILABLE
        else:
            ssrs_snapshot = SSRS_STATUS_STALE

        checks = run_startup_health_check(self.config_data, self.user_settings, ssrs_snapshot)
        logger.info("Startup health check: %r", checks)

        dlg = tk.Toplevel(self)
        dlg.title("Startup Health Check")
        dlg.transient(self)
        dlg.resizable(False, False)
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="System Readiness", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))
        for label, ok, detail in checks:
            row = ttk.Frame(frm)
            row.pack(fill="x", anchor="w", pady=2)
            mark = "✓" if ok else "✗"
            color = "#1a7f37" if ok else "#a05a00"
            ttk.Label(row, text=f"{mark} {label}", foreground=color, width=24, anchor="w").pack(side="left")
            if detail:
                ttk.Label(row, text=detail, foreground="#888").pack(side="left")
        ttk.Button(frm, text="OK", command=dlg.destroy).pack(anchor="e", pady=(10, 0))
        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

    def _show_settings_dialog(self, first_run=False):
        """Enhancements 2-4: one-time (or reopenable via 'Settings...')
        collection of the root save folder + operator profile."""
        dlg = tk.Toplevel(self)
        dlg.title("First-Time Setup" if first_run else "Settings")
        dlg.transient(self)
        dlg.resizable(False, False)

        pad = {"padx": 10, "pady": 6}
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)

        row = 0
        if first_run:
            ttk.Label(
                frm,
                text="Welcome! Set these up once - the app remembers them on every future launch.",
                wraplength=420, foreground="#444",
            ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 10))
            row += 1

        ttk.Label(frm, text="Root Save Folder:").grid(row=row, column=0, sticky="w", **pad)
        root_var = tk.StringVar(value=self.user_settings.get("root_save_folder", ""))
        ttk.Entry(frm, textvariable=root_var, width=42).grid(row=row, column=1, **pad)

        def browse():
            chosen = filedialog.askdirectory(title="Select Root Save Folder", parent=dlg)
            if chosen:
                root_var.set(chosen)

        ttk.Button(frm, text="Browse...", command=browse).grid(row=row, column=2, **pad)
        row += 1

        ttk.Label(frm, text="Operator Name:").grid(row=row, column=0, sticky="w", **pad)
        op_var = tk.StringVar(value=self.user_settings.get("operator_name", ""))
        ttk.Entry(frm, textvariable=op_var, width=42).grid(row=row, column=1, columnspan=2, sticky="w", **pad)
        row += 1

        ttk.Label(frm, text="Location:").grid(row=row, column=0, sticky="w", **pad)
        loc_var = tk.StringVar(value=self.user_settings.get("location", ""))
        ttk.Entry(frm, textvariable=loc_var, width=42).grid(row=row, column=1, columnspan=2, sticky="w", **pad)
        row += 1

        ttk.Label(frm, text="Business Unit:").grid(row=row, column=0, sticky="w", **pad)
        bu_var = tk.StringVar(value=self.user_settings.get("business_unit", ""))
        bu_names = [t["name"] for t in self._bu_templates_cache] if self._bu_templates_cache else []
        if bu_names:
            ttk.Combobox(frm, textvariable=bu_var, values=bu_names, state="readonly", width=39).grid(
                row=row, column=1, columnspan=2, sticky="w", **pad
            )
        else:
            ttk.Entry(frm, textvariable=bu_var, width=42).grid(row=row, column=1, columnspan=2, sticky="w", **pad)
        row += 1

        ttk.Label(frm, text="Email Address:").grid(row=row, column=0, sticky="w", **pad)
        email_var = tk.StringVar(value=self.user_settings.get("email", ""))
        ttk.Entry(frm, textvariable=email_var, width=42).grid(row=row, column=1, columnspan=2, sticky="w", **pad)
        row += 1

        admin_vars = {}
        if not first_run:
            # Admin / shared-install settings: SSRS Report URL, AD Lookup URL,
            # TracIT URL template, SSRS cache lifetime. These are saved to
            # admin_settings.json NEXT TO the app (not %APPDATA%), so unlike
            # the profile fields above, a change here applies to EVERYONE who
            # runs this same copy of the app - see _apply_admin_settings_overrides.
            ttk.Separator(frm, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=(12, 4))
            row += 1
            ttk.Label(
                frm, text="Admin / Shared Settings", font=("Segoe UI", 10, "bold"),
            ).grid(row=row, column=0, columnspan=3, sticky="w", padx=10)
            row += 1
            ttk.Label(
                frm,
                text="Applies to everyone using this installation - leave blank to keep the built-in default.",
                wraplength=420, foreground="#666",
            ).grid(row=row, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 6))
            row += 1

            admin_data = _load_admin_settings()
            admin_field_specs = [
                ("ssrs_report_url", "SSRS Report URL:", self.config_data.get("ssrs_asset_report", {}).get("url", "")),
                ("ad_lookup_url_template", "AD Lookup URL template:", self.config_data.get("ad_lookup_url_template", "")),
                ("tracit_url_template", "TracIT URL:", self.config_data.get("tracit_url_template", "")),
                ("ssrs_cache_max_age_hours", "SSRS cache max age (hours):",
                 str(self.config_data.get("ssrs_asset_report", {}).get("cache_max_age_hours", 24))),
            ]
            for key, label, current_value in admin_field_specs:
                ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w", **pad)
                v = tk.StringVar(value=admin_data.get(key, "") or "")
                entry = ttk.Entry(frm, textvariable=v, width=42)
                entry.grid(row=row, column=1, columnspan=2, sticky="w", **pad)
                # Show the effective (default or already-overridden) value as
                # placeholder-ish helper text below when the field is empty.
                if not v.get().strip():
                    ttk.Label(
                        frm, text=f"(current default: {current_value})", foreground="#888",
                    ).grid(row=row, column=1, columnspan=2, sticky="w", padx=(12, 0), pady=(24, 0))
                admin_vars[key] = v
                row += 1

        btn_row = ttk.Frame(frm)
        btn_row.grid(row=row, column=0, columnspan=3, pady=(10, 0))

        def do_save():
            if not root_var.get().strip():
                messagebox.showwarning("Root Save Folder required", "Please choose a root save folder.", parent=dlg)
                return
            self.user_settings = _save_user_settings({
                "root_save_folder": root_var.get().strip(),
                "operator_name": op_var.get().strip(),
                "location": loc_var.get().strip(),
                "business_unit": bu_var.get().strip(),
                "email": email_var.get().strip(),
            })
            if admin_vars:
                old_ssrs_url = self.config_data.get("ssrs_asset_report", {}).get("url", "")
                self.admin_settings = _save_admin_settings({k: v.get().strip() for k, v in admin_vars.items()})
                _apply_admin_settings_overrides(self.config_data, self.admin_settings)
                new_ssrs_url = self.config_data.get("ssrs_asset_report", {}).get("url", "")
                if hasattr(self, "_refresh_ssrs_status_labels"):
                    self._refresh_ssrs_status_labels()
                if new_ssrs_url != old_ssrs_url and hasattr(self, "_ssrs_refresh"):
                    # URL changed - the day-cache from the OLD url is no
                    # longer relevant, force a fresh download from the new one.
                    self._ssrs_refresh(force=True)
            if hasattr(self, "_refresh_save_path"):
                self._refresh_save_path()
            if hasattr(self, "_batch_refresh_save_path"):
                self._batch_refresh_save_path()
            self._refresh_header_operator_label()
            dlg.destroy()

        ttk.Button(btn_row, text="Save", command=do_save).pack(side="left", padx=6)
        if not first_run:
            ttk.Button(btn_row, text="Cancel", command=dlg.destroy).pack(side="left", padx=6)

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.grab_set()
        self.wait_window(dlg)

    def _make_scrollable(self, container):
        """Wraps `container` in a vertically-scrolling canvas and returns
        the inner content frame to build widgets into.

        Mousewheel/touchpad scrolling works anywhere over this area - not
        just when the cursor is directly over the canvas's own gutter (the
        old per-canvas <Enter>/<Leave> bind_all approach silently stopped
        working the moment the cursor was over any child widget, like an
        Entry or a Frame, since those swallow the Enter/Leave events before
        the canvas ever sees them). Instead every scrollable canvas
        registers itself once with the app-wide dispatcher set up in
        _init_global_mousewheel_scrolling(), which figures out - on every
        wheel/touchpad event, anywhere in the window - which registered
        canvas (if any) the cursor is currently over and scrolls that one."""
        outer = ttk.Frame(container)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        content = ttk.Frame(canvas, padding=16)
        window = canvas.create_window((0, 0), window=content, anchor="nw")

        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(window, width=e.width))

        self._register_scrollable_canvas(canvas)
        return content

    def _register_scrollable_canvas(self, canvas):
        if not hasattr(self, "_scrollable_canvases"):
            self._scrollable_canvases = []
            self._init_global_mousewheel_scrolling()
        self._scrollable_canvases.append(canvas)

    def _init_global_mousewheel_scrolling(self):
        """Binds mousewheel/touchpad scrolling ONCE for the whole window -
        every canvas registered via _register_scrollable_canvas() (Single
        Person, Bulk Batch, and any future scrollable panel) then scrolls
        correctly no matter where over that panel the cursor is, with no
        need to click the scrollbar itself first."""

        def _scroll_amount(event):
            # Windows/macOS send event.delta (multiples of 120 on Windows);
            # X11/Linux send Button-4 (up) / Button-5 (down) instead.
            if getattr(event, "num", None) == 4:
                return -1
            if getattr(event, "num", None) == 5:
                return 1
            delta = getattr(event, "delta", 0)
            if delta == 0:
                return 0
            return int(-1 * (delta / 120)) if abs(delta) >= 120 else (-1 if delta > 0 else 1)

        def _dispatch(event):
            amount = _scroll_amount(event)
            if amount == 0:
                return
            try:
                widget = self.winfo_containing(event.x_root, event.y_root)
            except Exception:
                widget = None
            while widget is not None:
                if widget in self._scrollable_canvases:
                    widget.yview_scroll(amount, "units")
                    return
                widget = widget.master

        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(seq, _dispatch)

    def _refresh_bu_combos(self):
        self._bu_templates_cache = list_templates(_resolve_path(self.config_data, "bu_templates_folder", "bu_templates"))
        names = [t["name"] for t in self._bu_templates_cache]
        logger.info("BU templates: loaded %d template(s): %r", len(names), names)
        if hasattr(self, "bu_combo"):
            self.bu_combo.config(values=names)
        if hasattr(self, "bulk_bu_combo"):
            self.bulk_bu_combo.config(values=names)

    def _add_bu_template(self):
        docx_path = filedialog.askopenfilename(
            title="Select the BU's base Word (.docx) form",
            filetypes=[("Word documents", "*.docx")],
        )
        if not docx_path:
            return
        bu_name = simpledialog.askstring("Business Unit name", "Enter a short name for this Business Unit:")
        if not bu_name or not bu_name.strip():
            return
        try:
            add_template(
                _resolve_path(self.config_data, "bu_templates_folder", "bu_templates"),
                bu_name.strip(), docx_path,
            )
        except Exception as exc:
            messagebox.showerror("Could not add template", str(exc))
            return
        messagebox.showinfo("Added", f"'{bu_name.strip()}' has been added. Select it from the dropdown.")
        self._refresh_bu_combos()
        self.bu_var.set(bu_name.strip())

    def _prompt_candidate_selection(self, candidates):
        result = {"choice": None}
        win = tk.Toplevel(self)
        win.title("Select the correct employee")
        win.grab_set()
        win.resizable(False, False)
        tk.Label(win, text=f"{len(candidates)} results found - please select the correct employee:",
                 font=("Segoe UI", 10, "bold")).pack(padx=15, pady=(15, 8))
        listbox = tk.Listbox(win, width=80, height=min(10, len(candidates)), font=("Segoe UI", 10))
        for c in candidates:
            summary = " | ".join(f"{k}: {v}" for k, v in c.items() if v)
            listbox.insert(tk.END, summary or "(no readable details)")
        listbox.pack(padx=15, pady=5)
        listbox.selection_set(0)

        def confirm():
            sel = listbox.curselection()
            if sel:
                result["choice"] = candidates[sel[0]]
            win.destroy()

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=(5, 15))
        tk.Button(btn_frame, text="Cancel", width=10, command=win.destroy).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Use Selected", width=14, command=confirm, bg="#2e7d32", fg="white").pack(side="left", padx=5)
        win.wait_window()
        return result["choice"]

    def _show_batch_preview(self, queue, warnings):
        """Shows everyone that was loaded from the Excel file in a table,
        so you can check it looks right BEFORE it's added to the Bulk
        Batch list. Returns True if the person clicked "Looks Good -
        Continue", False if they cancelled (nothing is loaded then)."""
        win = tk.Toplevel(self)
        win.title(f"Batch Preview - {len(queue)} people")
        win.geometry("760x520")
        win.grab_set()

        header = f"{len(queue)} people loaded. Review the list below before adding them."
        ttk.Label(win, text=header, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(12, 2))
        if warnings:
            ttk.Label(
                win, text=f"{len(warnings)} row(s) need attention - see the Note column below.",
                foreground="#a05a00",
            ).pack(anchor="w", padx=12, pady=(0, 6))

        columns = ("row", "emp_id", "type", "name", "manager", "serial", "note")
        headers = {"row": "Row", "emp_id": "Employee ID", "type": "Type", "name": "Name",
                   "manager": "Manager", "serial": "Laptop Serial", "note": "Note"}
        widths = {"row": 45, "emp_id": 100, "type": 110, "name": 150, "manager": 150, "serial": 110, "note": 190}

        tree_frame = ttk.Frame(win)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=6)
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=17)
        for c in columns:
            tree.heading(c, text=headers[c])
            tree.column(c, width=widths[c], anchor="w")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        for item in queue:
            notes = []
            if not item["type_recognized"]:
                notes.append("Type not recognized - will need manual fix")
            if not item.get("employee_name"):
                notes.append("Will look up via AD")
            else:
                notes.append("Using Excel data - no AD lookup")
            tree.insert("", "end", values=(
                item["row_number"], item["employee_id"], item["type"],
                item.get("employee_name") or "", item.get("manager_name") or "",
                item.get("laptop_serial_number") or "", "; ".join(notes),
            ))

        result = {"proceed": False}

        def start():
            result["proceed"] = True
            win.destroy()

        def cancel():
            win.destroy()

        btn_row = ttk.Frame(win)
        btn_row.pack(pady=(0, 12))
        ttk.Button(btn_row, text="Cancel", command=cancel).pack(side="left", padx=5)
        ttk.Button(btn_row, text="Looks Good - Continue", command=start).pack(side="left", padx=5)

        win.wait_window()
        return result["proceed"]

    # =====================================================================
    # SINGLE PERSON TAB
    # =====================================================================
    def _build_single_tab_shell(self):
        tab = self.single_tab

        bottom = ttk.Frame(tab, padding=(16, 8))
        bottom.pack(fill="x", side="bottom")

        # Enhancement: no manual Save-to/path/Choose Location on the main
        # screen anymore - the PDF always saves automatically to the root
        # folder configured once in Settings, under the automatic
        # <SubmissionType>/<Date>/<Operator>/ folder structure and
        # automatic <EmployeeID>_<Type>_<timestamp>.pdf filename (see
        # _compute_default_output_path). Where it landed is shown on the
        # success screen after Generate, with one-click Open PDF/Open Folder.
        self.output_path_override = None

        # Enhancement 10 - a compact Recent Documents panel (last 3), so
        # opening something just generated doesn't require a trip to the
        # separate Generated Documents tab.
        recent_frame = ttk.LabelFrame(bottom, text="Recent Documents", padding=(8, 4))
        recent_frame.pack(fill="x", pady=(0, 8))
        self.recent_docs_list_frame = ttk.Frame(recent_frame)
        self.recent_docs_list_frame.pack(fill="x")

        btn_row = ttk.Frame(bottom)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Generate Signed PDF", command=self._on_generate).pack(side="left")
        ttk.Button(btn_row, text="Start Another Form", command=self._on_start_another).pack(side="left", padx=10)

        self.single_content = self._make_scrollable(tab)
        self._recent_docs_refresh()

        self._sig_preview_image = {}
        self._signature_paths = {"employee_signature_path": None, "asset_receiver_signature_path": None}

    def _clear_single_content(self):
        for child in self.single_content.winfo_children():
            child.destroy()

    def _build_form(self):
        self._clear_single_content()
        parent = self.single_content

        self._build_bu_section(parent)
        ttk.Separator(parent).pack(fill="x", pady=8)
        self._build_employee_section(parent)
        ttk.Separator(parent).pack(fill="x", pady=8)
        self._build_contact_section(parent)
        ttk.Separator(parent).pack(fill="x", pady=8)
        self._build_submission_type_section(parent)
        self.asset_details_frame = ttk.Frame(parent)
        self.asset_details_frame.pack(fill="x", anchor="w", pady=(2, 0))
        self._rebuild_asset_details()
        ttk.Separator(parent).pack(fill="x", pady=8)

        sig_row = ttk.Frame(parent)
        sig_row.pack(fill="x", anchor="w")
        emp_sig_frame = ttk.Frame(sig_row)
        emp_sig_frame.pack(side="left", padx=(0, 40), anchor="n")
        self._build_signature_section(
            emp_sig_frame, "employee_signature_path", "Employee Signature", "Employee: sign here"
        )
        ar_sig_frame = ttk.Frame(sig_row)
        ar_sig_frame.pack(side="left", anchor="n")
        self._build_signature_section(
            ar_sig_frame, "asset_receiver_signature_path", "Asset Receiver Signature", "Asset Receiver: sign here"
        )

    # -------------------------------------------------------- 1: BU
    def _build_bu_section(self, parent):
        header = ttk.Frame(parent)
        header.pack(anchor="w", fill="x")
        ttk.Label(header, text="Business Unit", font=("Segoe UI", 11, "bold")).pack(side="left")
        self._add_info_tooltip(
            header,
            "Its Word template supplies the letterhead, address, and layout. Upload a BU's base "
            "Word (.docx) form once via 'Add New BU Template...' - the app automatically makes it "
            "fillable, no Acrobat or manual setup needed.",
        )

        templates = list_templates(_resolve_path(self.config_data, "bu_templates_folder", "bu_templates"))
        self._bu_templates_cache = templates

        row = ttk.Frame(parent)
        row.pack(anchor="w", fill="x", pady=(4, 0))
        self.bu_var = tk.StringVar(value="")
        names = [t["name"] for t in templates]
        self.bu_combo = ttk.Combobox(row, textvariable=self.bu_var, values=names, state="readonly", width=36)
        self.bu_combo.pack(side="left")
        if not names:
            self.status_var.set("No Business Unit templates yet - click 'Add New BU Template...'.")
        ttk.Button(row, text="Add New BU Template...", command=self._add_bu_template).pack(side="left", padx=8)

    # ----------------------------------------------------- 2: employee id
    def _build_employee_section(self, parent):
        header = ttk.Frame(parent)
        header.pack(anchor="w", fill="x")
        ttk.Label(header, text="Employee ID", font=("Segoe UI", 11, "bold")).pack(side="left")
        self._add_info_tooltip(
            header,
            "Scan or type the Employee ID, then press Enter to look it up automatically - SSRS is "
            "checked first (instant, no click needed); if it has no match, AD Lookup runs instead.",
        )

        row = ttk.Frame(parent)
        row.pack(anchor="w", pady=(4, 0))
        ttk.Label(row, text="Employee ID:").pack(side="left")
        self.emp_id_var = tk.StringVar(value="")
        digits_only_vcmd = (self.register(self._validate_digits_only), "%P")
        emp_id_entry = ttk.Entry(
            row, textvariable=self.emp_id_var, width=24,
            validate="key", validatecommand=digits_only_vcmd,
        )
        emp_id_entry.pack(side="left", padx=8)
        self.lookup_button = ttk.Button(row, text="Lookup (AD)", command=self._do_lookup)
        self.lookup_button.pack(side="left", padx=4)
        emp_id_entry.bind("<Return>", lambda e: self._do_lookup())

        # Enhancement 7 (relocated) - checked automatically the moment an
        # Employee ID is entered, right next to Lookup (AD), instead of a
        # separate checkbox buried in Asset Details further down. If SSRS
        # has a match it fills instantly (no network call - it's checked
        # against the already-downloaded/cached report); if not, this just
        # goes quiet and the normal Lookup (AD) button above still works
        # exactly as before.
        self.ssrs_check_label = ttk.Label(row, text="", foreground="#888")
        self.ssrs_check_label.pack(side="left", padx=(10, 0))
        self._ssrs_matched_record = None
        self._ssrs_check_after_id = None

        result_box = ttk.LabelFrame(parent, text="Employee Details", padding=8)
        result_box.pack(fill="x", pady=8)

        self.emp_name_var = tk.StringVar(value="")
        self.manager_name_var = tk.StringVar(value="")
        self._identity_lookup_source = "Manual"  # tracked for the audit log's Lookup Source column

        ttk.Label(result_box, text="Employee Name:").grid(row=0, column=0, sticky="w", pady=4)
        emp_name_entry = ttk.Entry(result_box, textvariable=self.emp_name_var, width=40)
        emp_name_entry.grid(row=0, column=1, sticky="w", pady=4, padx=6)
        self._bind_enter_advances_focus(emp_name_entry)

        ttk.Label(result_box, text="Manager Name:").grid(row=1, column=0, sticky="w", pady=4)
        manager_name_entry = ttk.Entry(result_box, textvariable=self.manager_name_var, width=40)
        manager_name_entry.grid(row=1, column=1, sticky="w", pady=4, padx=6)
        self._bind_enter_advances_focus(manager_name_entry)

        self.manual_entry_hint = ttk.Label(result_box, text="", foreground="#a05a00", wraplength=620)
        self.manual_entry_hint.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        def _refresh_save_path(*_a):
            # No visible "Save to" field anymore (it always auto-saves to
            # the Settings root folder + automatic subfolder/filename) -
            # kept as a callable no-op since other code still calls it
            # after things like a submission-type change.
            pass

        self._refresh_save_path = _refresh_save_path  # reused when submission type changes too
        self.emp_id_var.trace_add("write", _refresh_save_path)
        self.emp_name_var.trace_add("write", _refresh_save_path)
        self.emp_id_var.trace_add("write", self._schedule_ssrs_autocheck)

    def _schedule_ssrs_autocheck(self, *_args):
        """Debounced (600ms after typing stops) so this doesn't re-check on
        every single keystroke - only once the operator pauses or finishes
        typing/scanning the Employee ID."""
        if self._ssrs_check_after_id:
            try:
                self.after_cancel(self._ssrs_check_after_id)
            except Exception:
                pass
        self._ssrs_check_after_id = self.after(600, self._ssrs_autocheck_now)

    def _ssrs_autocheck_now(self):
        self._ssrs_check_after_id = None
        self._ssrs_matched_record = None
        emp_id = self.emp_id_var.get().strip()
        if not hasattr(self, "ssrs_check_label"):
            return
        if not emp_id or not emp_id.isdigit():
            self.ssrs_check_label.config(text="")
            self._ssrs_match_status = ""
            return
        index = self.ssrs_state.get("index", {})
        record = index.get(emp_id) or index.get("".join(c for c in emp_id if c.isdigit()))
        if not record:
            self._ssrs_match_status = "Not Found" if self.ssrs_state.get("status") == SSRS_STATUS_AVAILABLE else ""
            if self.ssrs_state.get("status") == SSRS_STATUS_AVAILABLE:
                self.ssrs_check_label.config(text="SSRS: no match - use Lookup (AD)", foreground="#a05a00")
            else:
                self.ssrs_check_label.config(text="")
            return
        self._ssrs_match_status = "Matched"
        self._ssrs_matched_record = record
        self._identity_lookup_source = "SSRS"
        self.ssrs_check_label.config(
            text=f"SSRS: matched - {record.get('employee_name', '') or emp_id}", foreground="#1a7f37",
        )
        if record.get("employee_name") and not self.emp_name_var.get().strip():
            self.emp_name_var.set(record["employee_name"])
            self.manual_entry_hint.config(text="Employee name auto-filled from SSRS Report.")
        self._apply_ssrs_record_to_asset_fields(record)
        # SSRS's column map has no manager field - if the manager name is
        # still blank after an SSRS match, quietly fall back to AD in the
        # background just for that (never overwrites a name SSRS already
        # supplied, and never pops up the multi-candidate picker for a
        # silent background fill).
        self._fetch_manager_via_ad_fallback(emp_id)

    def _fetch_manager_via_ad_fallback(self, emp_id):
        if self.manager_name_var.get().strip():
            return
        uses_webview = _uses_webview_lookup(self.config_data)

        def worker():
            try:
                if uses_webview:
                    candidates = self.browser_session.search(emp_id)
                else:
                    candidates = [fetch_employee_details(emp_id, self.config_data)]
                self.after(0, self._on_manager_fallback_result, candidates)
            except Exception:
                logger.debug("Manager AD fallback: lookup failed for %r (non-fatal)", emp_id, exc_info=True)

        threading.Thread(target=worker, daemon=True).start()

    def _on_manager_fallback_result(self, candidates):
        if not candidates or self.manager_name_var.get().strip():
            return
        manager = candidates[0].get("manager_name", "")
        if manager:
            self.manager_name_var.set(manager)

    def _apply_ssrs_record_to_asset_fields(self, record):
        """Fills the asset-detail fields (New Device Serial Number, Others)
        from a matched SSRS row - only meaningful for New Hire-type
        submissions, since that's the only case SSRS has issued-asset data
        for. Safe to call even if the asset-details widgets don't exist yet
        (e.g. no submission type picked yet) - it just no-ops via hasattr."""
        sub_type = self.submission_type_var.get() if hasattr(self, "submission_type_var") else ""
        if sub_type not in self.config_data.get("ssrs_asset_report", {}).get("new_hire_submission_types", []):
            return
        if record.get("serial_number") and hasattr(self, "new_serial_var"):
            self.new_serial_var.set(record["serial_number"])
        summary = self._format_ssrs_record(record)
        if summary and hasattr(self, "assets_other_var"):
            existing = self.assets_other_var.get().strip()
            self.assets_other_var.set(f"{existing}; {summary}" if existing else summary)
        if hasattr(self, "ssrs_autofill_hint"):
            self.ssrs_autofill_hint.config(text="Filled from SSRS.", foreground="#1a7f37")

    def _do_lookup(self):
        emp_id = self.emp_id_var.get().strip()
        if not emp_id:
            messagebox.showwarning("Missing Employee ID", "Please enter an Employee ID first.")
            return
        uses_webview = _uses_webview_lookup(self.config_data)
        self.status_var.set(f"Searching for {emp_id}...")
        self.lookup_button.config(state="disabled")

        def worker():
            try:
                if uses_webview:
                    candidates = self.browser_session.search(emp_id)
                else:
                    candidates = [fetch_employee_details(emp_id, self.config_data)]
                self.after(0, self._on_lookup_success, emp_id, candidates)
            except Exception as exc:
                self.after(0, self._on_lookup_failure, exc)

        threading.Thread(target=worker, daemon=True).start()

    def _on_lookup_success(self, emp_id, candidates):
        self.lookup_button.config(state="normal")
        if not candidates:
            self.status_var.set(f"No results for {emp_id} - enter details manually.")
            self.manual_entry_hint.config(text="No match found. Please enter the employee's name and manager manually below.")
            return
        if len(candidates) == 1:
            chosen = candidates[0]
        else:
            chosen = self._prompt_candidate_selection(candidates)
            if chosen is None:
                self.status_var.set("Selection cancelled.")
                return
        self.emp_name_var.set(chosen.get("emp_name", ""))
        self.manager_name_var.set(chosen.get("manager_name", ""))
        self._identity_lookup_source = "AD"
        self.manual_entry_hint.config(text="" if chosen.get("emp_name") else "Found a result, but no name field - enter manually.")
        self.status_var.set(f"Employee {emp_id} found.")

    def _on_lookup_failure(self, exc):
        # Friendly on-screen message - the real exception (connection
        # errors, missing config, etc.) goes to application.log only, not
        # in front of the operator.
        logger.error(
            "AD lookup failed for the active Employee ID (showing friendly message to operator): %s", exc,
            exc_info=(type(exc), exc, exc.__traceback__),
        )
        self.lookup_button.config(state="normal")
        self.status_var.set("Lookup unavailable - enter details manually.")
        self.manual_entry_hint.config(text="⚠ Employee lookup unavailable. Please enter details manually.")

    # ----------------------------------------------------- 3: phone
    def _build_contact_section(self, parent):
        ttk.Label(parent, text="Contact Number", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        row = ttk.Frame(parent)
        row.pack(anchor="w", pady=(4, 0))
        ttk.Label(row, text="Contact Number:").pack(side="left")
        self.contact_var = tk.StringVar(value="")
        contact_entry = ttk.Entry(row, textvariable=self.contact_var, width=24)
        contact_entry.pack(side="left", padx=8)
        self._bind_enter_advances_focus(contact_entry)

    # ------------------------------------------------- 4: submission type
    def _build_submission_type_section(self, parent):
        ttk.Label(parent, text="Submission Type", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        all_types = self.config_data.get("submission_type_row1", []) + self.config_data.get("submission_type_row2", [])
        self.submission_type_var = tk.StringVar(value="")

        box = ttk.Frame(parent)
        box.pack(anchor="w")
        for i, label in enumerate(all_types):
            ttk.Radiobutton(
                box, text=label, value=label, variable=self.submission_type_var,
                command=self._on_submission_type_changed,
            ).grid(row=i // 3, column=i % 3, sticky="w", padx=10, pady=4)

        self.submission_type_hint = ttk.Label(parent, text="", foreground="#a05a00")
        self.submission_type_hint.pack(anchor="w", pady=(4, 0))

    def _on_submission_type_changed(self):
        self.submission_type_hint.config(text="")
        self._rebuild_asset_details()
        if hasattr(self, "_refresh_save_path"):
            self._refresh_save_path()  # picks up the per-submission-type subfolder

    # ------------------------------------------------ 5: asset info
    def _rebuild_asset_details(self):
        for child in self.asset_details_frame.winfo_children():
            child.destroy()
        parent = self.asset_details_frame

        sub_type = self.submission_type_var.get()

        ttk.Label(parent, text="Asset Details", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(6, 2))

        self.current_serial_var = tk.StringVar(value="")
        self.new_serial_var = tk.StringVar(value="")
        self.assets_other_var = tk.StringVar(value="")
        self.asset_pending_var = tk.StringVar(value="")
        self.asset_checkbox_vars = {}

        if not sub_type:
            ttk.Label(parent, text="(select a submission type above)", foreground="#888").pack(anchor="w")
            return

        # The real production form always shows every asset field, for
        # every submission type that HAS an old/current device - a New
        # Hire has no existing device to return, so "Current Device
        # Serial Number" is hidden entirely for those types (see
        # no_current_asset_submission_types in CONFIG) rather than shown
        # blank and confusing.
        show_current_serial = sub_type not in self.config_data.get("no_current_asset_submission_types", [])
        if show_current_serial:
            row1 = ttk.Frame(parent)
            row1.pack(anchor="w", pady=4)
            ttk.Label(row1, text="Current Device Serial Number:").pack(side="left")
            ttk.Entry(row1, textvariable=self.current_serial_var, width=30).pack(side="left", padx=8)
            if sub_type in self.config_data.get("tracit_submission_types", []):
                ttk.Button(
                    row1, text="Open TracIT",
                    command=lambda: self._open_tracit(self.emp_id_var.get().strip(), self.current_serial_var.get().strip()),
                ).pack(side="left", padx=(4, 0))

        # Enhancement 7 - SSRS-based asset auto-population, for New Hire
        # (configurable via ssrs_asset_report.new_hire_submission_types).
        # The SSRS 'New SN' column is the device being ISSUED, so it fills
        # "New Device Serial Number" below - not "Current", which doesn't
        # apply to a New Hire at all (see show_current_serial above).
        # The rest of the SSRS row (hostname/model/ticket/monitor/dock/
        # etc.) is appended as a readable block into "Others (specify)",
        # since this app's Word templates don't have dedicated cells for
        # them yet - see the chat reply for how to get those added as
        # real fields instead.
        if sub_type in self.config_data.get("ssrs_asset_report", {}).get("new_hire_submission_types", []):
            ssrs_row = ttk.Frame(parent)
            ssrs_row.pack(anchor="w", pady=(0, 4))
            ttk.Label(ssrs_row, text="SSRS Report:", foreground="#666").pack(side="left")
            self.ssrs_autofill_hint = ttk.Label(ssrs_row, text="", foreground="#888")
            self.ssrs_autofill_hint.pack(side="left", padx=(6, 0))
            # Auto-fill is checked automatically near the Employee ID field
            # above the moment the ID is entered (see _ssrs_autocheck_now) -
            # if a match was already found for the current Employee ID,
            # re-apply it here too since these fields were just rebuilt.
            if self._ssrs_matched_record:
                self._apply_ssrs_record_to_asset_fields(self._ssrs_matched_record)

        row2 = ttk.Frame(parent)
        row2.pack(anchor="w", pady=4)
        ttk.Label(row2, text="New Device Serial Number:").pack(side="left")
        ttk.Entry(row2, textvariable=self.new_serial_var, width=30).pack(side="left", padx=8)

        ttk.Label(parent, text="List of Assets Issued/Return:").pack(anchor="w", pady=(14, 2))
        checklist_frame = ttk.Frame(parent)
        checklist_frame.pack(anchor="w")
        items = self.config_data.get("asset_checklist_items", [])
        for i, item in enumerate(items):
            var = tk.BooleanVar()
            self.asset_checkbox_vars[item] = var
            ttk.Checkbutton(checklist_frame, text=item, variable=var).grid(
                row=i // 3, column=i % 3, sticky="w", padx=8, pady=3
            )

        row3 = ttk.Frame(parent)
        row3.pack(anchor="w", pady=(10, 4))
        ttk.Label(row3, text="Others (specify):").pack(side="left")
        ttk.Entry(row3, textvariable=self.assets_other_var, width=40).pack(side="left", padx=8)

        row4 = ttk.Frame(parent)
        row4.pack(anchor="w", pady=4)
        ttk.Label(row4, text="Asset Pending for Submission (if any):").pack(side="left")
        ttk.Entry(row4, textvariable=self.asset_pending_var, width=40).pack(side="left", padx=8)

    @staticmethod
    def _format_ssrs_record(record):
        """Turns a parsed SSRS asset row into a COMPACT one-line summary
        (e.g. '✅ Device Model: Latitude 5420 | Ticket: INC0012345 | ...'),
        appended into the 'Others (specify)' free-text field since the Word
        templates don't have dedicated cells for these fields yet. Device
        Model and Ticket Number lead since they're what's most often needed
        at a glance; the rest follow only if the report actually has them."""
        labels = [
            ("device_model", "Device Model"), ("ticket_number", "Ticket"),
            ("hostname", "Hostname"), ("monitor", "Monitor"),
            ("dock", "Dock"), ("keyboard_mouse", "Keyboard & Mouse"),
            ("power_adapter", "Power Adapter"), ("battery", "Battery"),
        ]
        parts = [f"{label}: {record.get(key, '')}" for key, label in labels if record.get(key)]
        if not parts:
            return ""
        return "✅ " + " | ".join(parts)

    # ---------------------------------------------- 6/7: signatures
    def _build_signature_section(self, parent, data_key, title, capture_title):
        ttk.Label(parent, text=title, font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(
            parent,
            text="Captured by this app (Topaz pad if enabled, otherwise on-screen) and "
            "cryptographically signed into the final PDF - no Adobe Acrobat needed.",
            wraplength=320, foreground="#555",
        ).pack(anchor="w", pady=(4, 10))

        preview_label = ttk.Label(parent, text="(not signed yet)", relief="groove", width=40, anchor="center")
        preview_label.pack(pady=(0, 10))
        setattr(self, f"_preview_label_{data_key}", preview_label)

        def show_preview(path):
            try:
                img = Image.open(path)
                img.load()
                # Topaz pads classically write 1-bit monochrome or
                # palette bitmaps (ImageFileFormat/JustifyMode etc. don't
                # control this) - ImageTk.PhotoImage can render those as
                # blank/black instead of raising, which would make the
                # preview look "not showing" with no error anywhere.
                # Normalizing to RGB first is cheap and makes the
                # thumbnail render correctly regardless of the source mode.
                if img.mode not in ("RGB", "RGBA"):
                    img = img.convert("RGB")
                img.thumbnail((280, 110))
                photo = ImageTk.PhotoImage(img)
                self._sig_preview_image[data_key] = photo  # keep a reference alive
                preview_label.image = photo  # belt-and-suspenders against GC
                preview_label.config(image=photo, text="", compound="image")
            except Exception:
                logger.exception("Signature preview: could not render thumbnail for %s (%s)", data_key, path)
                preview_label.config(text="(signature captured)")

        existing_path = self._signature_paths.get(data_key)
        if existing_path and os.path.exists(existing_path):
            show_preview(existing_path)

        def sign_now():
            path = self.signature_service.capture(self, title=capture_title)
            if path:
                self._signature_paths[data_key] = path
                show_preview(path)
                self.status_var.set(f"{title.split('. ', 1)[-1]} captured.")

        ttk.Button(parent, text="Sign Now...", command=sign_now).pack()

    def _reset_signature_previews(self):
        for key in ("employee_signature_path", "asset_receiver_signature_path"):
            self._signature_paths[key] = None
            self._sig_preview_image.pop(key, None)
            label = getattr(self, f"_preview_label_{key}", None)
            if label is not None:
                label.config(image="", text="(not signed yet)")

    # ------------------------------------------------- save location
    def _compute_default_output_path(self):
        # Enhancement 5 folder structure + Enhancement 6 filename
        # convention: <Root>/<SubmissionType>/<Date>/<Operator>/
        # <EmployeeID>_<SubmissionType>_<YYYYMMDD_HHMMSS>.pdf
        # Root comes from the one-time Settings dialog (Enhancements 2-4)
        # once configured; falls back to CONFIG's pdf_save_folder until
        # then, so nothing breaks for anyone who hasn't set it up yet.
        root = (self.user_settings.get("root_save_folder") or "").strip() or _resolve_path(
            self.config_data, "pdf_save_folder", "generated_forms"
        )
        emp_id = self.emp_id_var.get().strip() if hasattr(self, "emp_id_var") else ""
        sub_type = self.submission_type_var.get().strip() if hasattr(self, "submission_type_var") else ""
        operator = self.user_settings.get("operator_name") or "Unspecified_Operator"
        today_str = date.today().strftime("%d-%b-%Y")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        compact_type = SUBMISSION_TYPE_TAGS.get(sub_type) or _safe_folder_name(sub_type)
        filename = f"{emp_id or 'unknown'}_{compact_type}_{timestamp}.pdf"
        return os.path.join(
            root, _safe_folder_name(sub_type), today_str, _safe_folder_name(operator), filename
        )

    @staticmethod
    def _validate_digits_only(proposed_value):
        """Tkinter entry validatecommand: only digits (or an empty box, so
        backspace/select-all-delete still work) are allowed in the Employee
        ID field - keystrokes and pastes of anything else are rejected
        outright rather than accepted and then complained about."""
        return proposed_value == "" or proposed_value.isdigit()

    def _open_tracit(self, employee_id, serial_number):
        """Opens the configured TracIT page. If tracit_url_template
        contains {employee_id}/{serial_number} placeholders (once a real
        deep-link URL format is confirmed), they're filled in; otherwise
        the plain base page opens as-is, same as clicking a bookmark."""
        template = self.config_data.get("tracit_url_template", "").strip()
        if not template:
            messagebox.showwarning(
                "TracIT not configured",
                "tracit_url_template isn't set in CONFIG yet.",
            )
            return
        try:
            url = template.format(employee_id=employee_id or "", serial_number=serial_number or "")
        except Exception:
            logger.debug("TracIT: URL template has no/invalid placeholders, opening as-is", exc_info=True)
            url = template
        logger.info("TracIT: opening %s (employee_id=%r, serial_number=%r)", url, employee_id, serial_number)
        webbrowser.open(url)

    # ------------------------------------------------------------ generate
    def _validate_form(self):
        if not self.bu_var.get():
            messagebox.showwarning("Business Unit required", "Please select a Business Unit.")
            return False
        if not self.emp_id_var.get().strip():
            messagebox.showwarning("Missing Employee ID", "Please enter an Employee ID.")
            return False
        if not self.emp_id_var.get().strip().isdigit():
            # Belt-and-suspenders alongside the digits-only entry
            # validation (registered in _build_employee_section) - this
            # catches anything that slips past keystroke validation, e.g.
            # a value set programmatically.
            messagebox.showwarning("Invalid Employee ID", "Employee ID must contain numbers only.")
            return False
        if not self.emp_name_var.get().strip():
            messagebox.showwarning(
                "Missing Employee Name",
                "Please look up the employee, or enter their name manually, before continuing.",
            )
            return False
        if not self.contact_var.get().strip():
            messagebox.showwarning("Missing contact number", "Please enter a contact number.")
            return False
        if not self.submission_type_var.get():
            messagebox.showwarning("Missing selection", "Please select a submission type.")
            return False
        if not self._signature_paths.get("employee_signature_path"):
            messagebox.showwarning("Signature required", "Please capture the Employee Signature.")
            return False
        if not self._signature_paths.get("asset_receiver_signature_path"):
            messagebox.showwarning("Signature required", "Please capture the Asset Receiver Signature.")
            return False
        return True

    def _on_generate(self):
        logger.info(
            "Generate clicked: emp_id=%r bu=%r submission_type=%r",
            self.emp_id_var.get().strip(), self.bu_var.get(), self.submission_type_var.get(),
        )
        if not self._validate_form():
            logger.info("Generate: form validation failed, stopping")
            return

        chosen_bu = next((t for t in self._bu_templates_cache if t["name"] == self.bu_var.get()), None)
        if not chosen_bu:
            logger.error("Generate: chosen Business Unit %r not found in template cache", self.bu_var.get())
            messagebox.showerror("Business Unit error", "The selected Business Unit template could not be found - pick it again.")
            return
        bu_fillable_docx_path = chosen_bu["fillable_path"]

        today_str = date.today().strftime("%d-%b-%Y")
        emp_id = self.emp_id_var.get().strip()
        emp_name = self.emp_name_var.get().strip()
        manager_name = self.manager_name_var.get().strip()
        contact_number = self.contact_var.get().strip()
        chosen_type = self.submission_type_var.get()

        field_values = {
            "Date": today_str, "EmpID": emp_id, "EmpName": emp_name,
            "ContactNumber": contact_number, "ManagerName": manager_name,
            "LastWorkingDate": "",
        }

        checked_tags = {tag for label, tag in SUBMISSION_TYPE_TAGS.items() if label == chosen_type}

        current_serial = self.current_serial_var.get().strip()
        new_serial = self.new_serial_var.get().strip()
        assets_other = self.assets_other_var.get().strip()
        asset_pending = self.asset_pending_var.get().strip()
        assets_issued = [k for k, v in self.asset_checkbox_vars.items() if v.get()]

        field_values["CurrentSerial"] = current_serial
        field_values["NewSerial"] = new_serial
        field_values["Others"] = assets_other
        field_values["AssetPending"] = asset_pending

        checked_tags |= {tag for label, tag in ASSET_CHECKLIST_TAGS.items() if label in assets_issued}

        signature_images = {
            "EmployeeSignature": self._signature_paths.get("employee_signature_path"),
            "AssetReceiverSignature": self._signature_paths.get("asset_receiver_signature_path"),
        }

        output_path = self.output_path_override or self._compute_default_output_path()
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        operator_name = self.user_settings.get("operator_name") or ""
        operator_location = self.user_settings.get("location") or ""
        ssrs_match_status = getattr(self, "_ssrs_match_status", "") or ""

        tmp_dir = tempfile.mkdtemp(prefix="itasset_")
        filled_docx = os.path.join(tmp_dir, "filled.docx")
        unsigned_pdf = os.path.join(tmp_dir, "unsigned.pdf")

        self.status_var.set("Filling document...")
        self.update_idletasks()
        try:
            fill_docx(bu_fillable_docx_path, filled_docx, field_values, checked_tags, signature_images)

            self.status_var.set("Exporting to PDF via Word...")
            self.update_idletasks()
            convert_docx_to_pdf_via_word(filled_docx, unsigned_pdf)

            # Enhancement 13/18 - visible "Signed On: <timestamp>" caption
            # under both signatures. Also must happen BEFORE signing.
            _stamp_signed_on_captions(unsigned_pdf, _find_signature_field_boxes(unsigned_pdf))

            # Enhancement 19 - PDF metadata. Must happen BEFORE signing:
            # editing bytes after the cryptographic signature is applied
            # would invalidate it.
            _set_pdf_metadata(unsigned_pdf, {
                "Title": f"IT Asset Submission Acknowledgement - {emp_name}",
                "Author": operator_name or "IT Asset Submission Acknowledgement System",
                "Subject": chosen_type,
                "EmployeeName": emp_name, "EmployeeID": emp_id, "ManagerName": manager_name,
                "SubmissionType": chosen_type, "GeneratedBy": operator_name,
                "GeneratedDate": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ProbTicket": "", "SCID": "",
            })

            self.status_var.set("Applying cryptographic signatures...")
            self.update_idletasks()
            # Enhancement 12 - Adobe's own Signature Properties dialog has
            # no dedicated Employee ID field, so it's folded into the
            # Reason text (along with the employee name) for both
            # signatures, alongside the real Location and the automatic
            # signing timestamp pyHanko already embeds.
            base_reason = self.config_data.get("signing_reason") or "IT Asset Acknowledgement"
            emp_reason = f"{base_reason} - {emp_name} (Employee ID: {emp_id})"
            sign_pdf_with_signatures(
                unsigned_pdf, output_path,
                [
                    {"field_name": "EmployeeSignature", "display_name": emp_name,
                     "location": operator_location, "reason": emp_reason},
                    {"field_name": "AssetReceiverSignature",
                     "display_name": self.config_data.get("asset_receiver_display_name", "Asset Receiver"),
                     "location": operator_location, "reason": emp_reason},
                ],
                self.config_data,
            )
        except (FillDocxError, WordConversionError, SigningError) as exc:
            logger.error("Generate failed for emp_id=%r: %s", emp_id, exc)
            messagebox.showerror("Generation Error", str(exc))
            self.status_var.set("Generation failed.")
            _write_audit_log_entry(
                operator=operator_name, employee_id=emp_id, submission_type=chosen_type,
                ssrs_match_status=ssrs_match_status, lookup_source=getattr(self, "_identity_lookup_source", ""),
                pdf_generated=False, pdf_location="", email_draft_created=False,
            )
            return
        except Exception as exc:
            logger.exception("Generate failed unexpectedly for emp_id=%r", emp_id)
            messagebox.showerror("Generation Error", f"Unexpected error: {exc}")
            self.status_var.set("Generation failed.")
            _write_audit_log_entry(
                operator=operator_name, employee_id=emp_id, submission_type=chosen_type,
                ssrs_match_status=ssrs_match_status, lookup_source=getattr(self, "_identity_lookup_source", ""),
                pdf_generated=False, pdf_location="", email_draft_created=False,
            )
            return
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

        logger.info("Generate SUCCESS for emp_id=%r -> %s", emp_id, output_path)

        self.status_var.set(f"Saved (signed): {output_path}")

        email_draft_created = False
        if _is_email_draft_type(self.config_data, chosen_type):
            email_data = {
                "emp_id": emp_id, "emp_name": emp_name, "manager_name": manager_name,
                "submission_type": chosen_type, "date": today_str,
                "current_device_serial": current_serial, "assets_issued": assets_issued,
                "assets_other": assets_other, "asset_pending": asset_pending,
                "email_action_label": self.config_data.get("email_templates", {}).get(chosen_type, chosen_type),
            }
            try:
                open_draft(email_data, to_addresses=self.config_data.get("lwd_email_to", ""), attachment_path=output_path)
                email_draft_created = True
            except RuntimeError as exc:
                messagebox.showwarning("Email draft not opened", str(exc))

        _write_audit_log_entry(
            operator=operator_name, employee_id=emp_id, submission_type=chosen_type,
            ssrs_match_status=ssrs_match_status, lookup_source=getattr(self, "_identity_lookup_source", ""),
            pdf_generated=True, pdf_location=output_path, email_draft_created=email_draft_created,
        )
        if hasattr(self, "_history_refresh"):
            self._history_refresh()
        if hasattr(self, "_recent_docs_refresh"):
            self._recent_docs_refresh()

        self._show_generation_success_dialog(output_path)

    def _show_generation_success_dialog(self, output_path):
        """Enhancement 9 - a compact success screen (instead of a plain
        OK-only messagebox) with the two things an operator does next
        already one click away, plus a fast way to move on to the next
        person without any extra navigation."""
        dlg = tk.Toplevel(self)
        dlg.title("PDF Generated")
        dlg.resizable(False, False)
        dlg.transient(self)

        frm = ttk.Frame(dlg, padding=18)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="✅ PDF Generated Successfully", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(
            frm, text=output_path, foreground="#555", wraplength=440,
        ).pack(anchor="w", pady=(4, 2))
        ttk.Label(
            frm,
            text="Filled and cryptographically signed - nothing further to do in Acrobat.",
            foreground="#888", wraplength=440,
        ).pack(anchor="w", pady=(0, 14))

        btn_row = ttk.Frame(frm)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Open PDF", command=lambda: self._open_with_os_default(output_path)).pack(
            side="left"
        )
        ttk.Button(
            btn_row, text="Open Folder",
            command=lambda: self._open_with_os_default(os.path.dirname(output_path)),
        ).pack(side="left", padx=8)

        def start_next():
            dlg.destroy()
            self._on_start_another()

        ttk.Button(btn_row, text="Start Next Employee", command=start_next).pack(side="left", padx=8)
        ttk.Button(frm, text="Close", command=dlg.destroy).pack(anchor="e", pady=(14, 0))

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.grab_set()

    def _reset_person_fields(self, keep_bu=True):
        if not keep_bu:
            self.bu_var.set("")
        self.emp_id_var.set("")
        self.emp_name_var.set("")
        self.manager_name_var.set("")
        self.manual_entry_hint.config(text="")
        self.contact_var.set("")
        self.submission_type_var.set("")
        self.submission_type_hint.config(text="")
        self._rebuild_asset_details()
        self._reset_signature_previews()
        self.output_path_override = None
        self._identity_lookup_source = "Manual"
        self.status_var.set("Ready.")

    def _on_start_another(self):
        self._reset_person_fields(keep_bu=True)

    # =====================================================================
    # BULK BATCH TAB
    # =====================================================================
    def _build_bulk_tab_shell(self):
        tab = self.bulk_tab
        self.batch_queue = []
        self._batch_active_index = None
        self._batch_auto_advance = False

        ttk.Label(tab, text="Bulk Batch Processing", font=("Segoe UI", 13, "bold")).pack(
            anchor="w", padx=16, pady=(14, 2)
        )
        ttk.Label(
            tab,
            text="Import an Excel list of people, then process them one at a time - in any order you like. "
            "Each row's status updates live as you go.",
            foreground="#666", wraplength=880,
        ).pack(anchor="w", padx=16, pady=(0, 10))

        toolbar = ttk.Frame(tab, padding=(16, 0))
        toolbar.pack(fill="x")
        ttk.Label(toolbar, text="Business Unit (default for this batch):").pack(side="left")
        self.bulk_bu_var = tk.StringVar(value="")
        self.bulk_bu_combo = ttk.Combobox(
            toolbar, textvariable=self.bulk_bu_var,
            values=[t["name"] for t in self._bu_templates_cache], state="readonly", width=30,
        )
        self.bulk_bu_combo.pack(side="left", padx=8)
        self.bulk_bu_combo.bind("<<ComboboxSelected>>", self._batch_on_default_bu_changed)
        ttk.Button(toolbar, text="Add New BU Template...", command=self._add_bu_template).pack(side="left", padx=(10, 0))
        ttk.Label(
            toolbar, text="(a person's own 'Business Unit' column in the Excel file overrides this)",
            foreground="#888",
        ).pack(side="left", padx=10)

        import_row = ttk.Frame(tab, padding=(16, 10, 16, 4))
        import_row.pack(fill="x")
        ttk.Button(import_row, text="Import Batch from Excel...", command=self._batch_import_excel).pack(side="left")
        ttk.Button(import_row, text="Download Excel Template...", command=self._batch_download_template).pack(
            side="left", padx=10
        )
        ttk.Label(
            import_row, text="New to this? Download the template, fill it in, then import it.",
            foreground="#555",
        ).pack(side="left", padx=10)

        summary_row = ttk.Frame(tab, padding=(16, 4, 16, 8))
        summary_row.pack(fill="x")
        self.batch_summary_var = tk.StringVar(value="No batch loaded yet.")
        ttk.Label(summary_row, textvariable=self.batch_summary_var, font=("Segoe UI", 10, "bold")).pack(side="left")
        self.process_all_button = ttk.Button(summary_row, text="Process All (auto-advance)", command=self._batch_process_all)
        self.process_all_button.pack(side="right")
        self.stop_auto_button = ttk.Button(summary_row, text="Stop Auto-Advance", command=self._batch_stop_auto)

        list_frame = ttk.Frame(tab, padding=(16, 0, 16, 4))
        list_frame.pack(fill="both", expand=False)
        columns = ("row", "emp_id", "name", "manager", "type", "serial", "status")
        headers = {"row": "Row", "emp_id": "Employee ID", "name": "Name", "manager": "Manager",
                   "type": "Type", "serial": "Laptop Serial", "status": "Status"}
        widths = {"row": 45, "emp_id": 100, "name": 160, "manager": 150, "type": 110, "serial": 110, "status": 130}
        self.batch_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=9)
        for c in columns:
            self.batch_tree.heading(c, text=headers[c])
            self.batch_tree.column(c, width=widths[c], anchor="w")
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.batch_tree.yview)
        self.batch_tree.configure(yscrollcommand=vsb.set)
        self.batch_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.batch_tree.bind("<Double-1>", lambda e: self._batch_process_selected())
        self.batch_tree.tag_configure("completed", background="#dff0d8")
        self.batch_tree.tag_configure("skipped", background="#eeeeee", foreground="#777777")
        self.batch_tree.tag_configure("attention", background="#fff3cd")
        self.batch_tree.tag_configure("pending", background="white")
        self.batch_tree.tag_configure("active", background="#cfe2ff")

        row_btns = ttk.Frame(tab, padding=(16, 0, 16, 8))
        row_btns.pack(fill="x")
        ttk.Button(row_btns, text="Process Selected", command=self._batch_process_selected).pack(side="left")
        ttk.Button(row_btns, text="Skip Selected", command=self._batch_skip_selected).pack(side="left", padx=10)
        ttk.Button(row_btns, text="Requeue Selected", command=self._batch_requeue_selected).pack(side="left")
        ttk.Label(
            row_btns,
            text="Tip: double-click any row to process that person now - or Skip and come back to them later.",
            foreground="#888",
        ).pack(side="left", padx=14)

        ttk.Separator(tab).pack(fill="x", padx=16, pady=(4, 0))

        panel_container = ttk.Frame(tab)
        panel_container.pack(fill="both", expand=True)
        self.batch_panel_content = self._make_scrollable(panel_container)

        self.batch_placeholder = ttk.Label(
            self.batch_panel_content,
            text="Select a person from the list above (double-click a row), or click "
            "'Process All (auto-advance)' to begin.",
            foreground="#888",
        )
        self.batch_placeholder.pack(anchor="w", pady=20)
        self.batch_person_frame = ttk.Frame(self.batch_panel_content)

        self._batch_refresh_tree()

    # =====================================================================
    # GENERATED DOCUMENTS TAB (Enhancement 15) - a read-only, searchable
    # view over the audit log's successful-generation rows, with quick
    # actions to open the PDF or its containing folder.
    # =====================================================================
    def _build_history_tab_shell(self):
        tab = self.history_tab
        ttk.Label(tab, text="Generated Documents", font=("Segoe UI", 13, "bold")).pack(
            anchor="w", padx=16, pady=(14, 2)
        )
        ttk.Label(
            tab,
            text="Every PDF this app has generated on this machine, newest first.",
            foreground="#666",
        ).pack(anchor="w", padx=16, pady=(0, 10))

        toolbar = ttk.Frame(tab)
        toolbar.pack(fill="x", padx=16, pady=(0, 8))
        ttk.Label(toolbar, text="Search:").pack(side="left")
        self.history_search_var = tk.StringVar(value="")
        search_entry = ttk.Entry(toolbar, textvariable=self.history_search_var, width=30)
        search_entry.pack(side="left", padx=6)
        search_entry.bind("<KeyRelease>", lambda e: self._history_refresh())
        ttk.Button(toolbar, text="Refresh", command=self._history_refresh).pack(side="left", padx=(6, 0))

        columns = ("employee_id", "submission_type", "timestamp", "pdf_location")
        self.history_tree = ttk.Treeview(tab, columns=columns, show="headings", height=16)
        self.history_tree.heading("employee_id", text="Employee ID")
        self.history_tree.heading("submission_type", text="Submission Type")
        self.history_tree.heading("timestamp", text="Generated Date")
        self.history_tree.heading("pdf_location", text="PDF Location")
        self.history_tree.column("employee_id", width=110, anchor="w")
        self.history_tree.column("submission_type", width=140, anchor="w")
        self.history_tree.column("timestamp", width=150, anchor="w")
        self.history_tree.column("pdf_location", width=420, anchor="w")
        self.history_tree.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        self._register_scrollable_canvas(self.history_tree)  # same global mousewheel dispatcher

        action_row = ttk.Frame(tab)
        action_row.pack(anchor="w", padx=16, pady=(0, 14))
        ttk.Button(action_row, text="Open PDF", command=self._history_open_pdf).pack(side="left")
        ttk.Button(action_row, text="Open Folder", command=self._history_open_folder).pack(side="left", padx=8)

        self._history_refresh()

    def _recent_docs_refresh(self):
        """Enhancement 10 - the compact panel on the Single Person tab
        (last 3 generated PDFs), separate from the full searchable
        Generated Documents tab."""
        if not hasattr(self, "recent_docs_list_frame"):
            return
        for child in self.recent_docs_list_frame.winfo_children():
            child.destroy()
        entries = [e for e in _read_audit_log_entries() if (e.get("pdf_generated") or "").strip().lower() == "yes"][:3]
        if not entries:
            ttk.Label(self.recent_docs_list_frame, text="No documents generated yet.", foreground="#888").pack(
                anchor="w"
            )
            return
        for entry in entries:
            path = entry.get("pdf_location", "")
            row = ttk.Frame(self.recent_docs_list_frame)
            row.pack(fill="x", pady=1)
            ttk.Label(
                row,
                text=f"{entry.get('employee_id', '')}  •  {entry.get('submission_type', '')}  •  {entry.get('timestamp', '')}",
                anchor="w",
            ).pack(side="left", fill="x", expand=True)

            def _open_recent_pdf(p=path):
                if p and os.path.exists(p):
                    self._open_with_os_default(p)
                else:
                    messagebox.showwarning("File not found", f"This PDF no longer exists at:\n{p}")

            def _open_recent_folder(p=path):
                folder = os.path.dirname(p)
                if folder and os.path.isdir(folder):
                    self._open_with_os_default(folder)
                else:
                    messagebox.showwarning("Folder not found", f"This folder no longer exists:\n{folder}")

            ttk.Button(row, text="Open", width=6, command=_open_recent_pdf).pack(side="left", padx=(4, 2))
            ttk.Button(row, text="Folder", width=7, command=_open_recent_folder).pack(side="left")

    def _history_refresh(self):
        if not hasattr(self, "history_tree"):
            return
        for row_id in self.history_tree.get_children():
            self.history_tree.delete(row_id)
        query = (self.history_search_var.get() if hasattr(self, "history_search_var") else "").strip().lower()
        for entry in _read_audit_log_entries():
            if (entry.get("pdf_generated") or "").strip().lower() != "yes":
                continue
            haystack = " ".join([
                entry.get("employee_id", ""), entry.get("submission_type", ""), entry.get("pdf_location", ""),
            ]).lower()
            if query and query not in haystack:
                continue
            self.history_tree.insert(
                "", "end",
                values=(
                    entry.get("employee_id", ""), entry.get("submission_type", ""),
                    entry.get("timestamp", ""), entry.get("pdf_location", ""),
                ),
            )

    def _history_selected_pdf_path(self):
        selection = self.history_tree.selection()
        if not selection:
            messagebox.showinfo("No row selected", "Select a row in the list first.")
            return None
        values = self.history_tree.item(selection[0], "values")
        return values[3] if len(values) > 3 else None

    def _history_open_pdf(self):
        path = self._history_selected_pdf_path()
        if not path:
            return
        if not os.path.exists(path):
            messagebox.showwarning("File not found", f"This PDF no longer exists at:\n{path}")
            return
        self._open_with_os_default(path)

    def _history_open_folder(self):
        path = self._history_selected_pdf_path()
        if not path:
            return
        folder = os.path.dirname(path)
        if not os.path.isdir(folder):
            messagebox.showwarning("Folder not found", f"This folder no longer exists:\n{folder}")
            return
        self._open_with_os_default(folder)

    @staticmethod
    def _open_with_os_default(path):
        try:
            if sys.platform == "win32":
                os.startfile(path)  # noqa: S606 - intentional, opens in Explorer/the default PDF viewer
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception:
            logger.exception("Generated Documents: could not open %r", path)
            messagebox.showerror("Could not open", f"Could not open:\n{path}")

    # ---------------------------------------------------------- import
    def _batch_import_excel(self):
        xlsx_path = filedialog.askopenfilename(
            title="Select the batch Excel file",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not xlsx_path:
            return
        try:
            queue, warnings = load_queue(xlsx_path, self.config_data.get("bulk_type_aliases", {}))
        except BulkQueueError as exc:
            messagebox.showerror("Could not load batch file", str(exc))
            return

        if not self._show_batch_preview(queue, warnings):
            return  # cancelled - nothing added

        for item in queue:
            if not item["type_recognized"]:
                item["status"] = "Needs Attention"

        self.batch_queue = queue
        self._batch_active_index = None
        self._batch_auto_advance = False
        self._batch_clear_panel()
        self._batch_refresh_tree()
        messagebox.showinfo(
            "Batch loaded",
            f"{len(queue)} people added to the list below. Pick a Business Unit above, then click "
            "'Process All' or double-click any row to start with that person.",
        )

    def _batch_download_template(self):
        path = filedialog.asksaveasfilename(
            title="Save Excel Template As",
            initialfile="IT_Asset_Batch_Template.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        try:
            create_batch_template_xlsx(path, self.config_data)
        except Exception as exc:
            messagebox.showerror("Could not create template", str(exc))
            return
        messagebox.showinfo(
            "Template saved",
            f"Saved to:\n{path}\n\nFill it in (see the 'Instructions' sheet) and use "
            "'Import Batch from Excel...' to load it.",
        )

    # ---------------------------------------------------------- tree/list
    def _batch_refresh_tree(self):
        self.batch_tree.delete(*self.batch_tree.get_children())
        for i, item in enumerate(self.batch_queue):
            tag = {
                "Completed": "completed", "Skipped": "skipped",
                "Needs Attention": "attention", "Pending": "pending",
            }.get(item["status"], "pending")
            if i == self._batch_active_index:
                tag = "active"
            self.batch_tree.insert("", "end", iid=str(i), values=(
                item["row_number"], item["employee_id"], item.get("employee_name") or "",
                item.get("manager_name") or "", item.get("type") or "",
                item.get("laptop_serial_number") or "", item["status"],
            ), tags=(tag,))

        total = len(self.batch_queue)
        done = sum(1 for it in self.batch_queue if it["status"] == "Completed")
        skipped = sum(1 for it in self.batch_queue if it["status"] == "Skipped")
        attention = sum(1 for it in self.batch_queue if it["status"] == "Needs Attention")
        pending = total - done - skipped - attention
        if total:
            self.batch_summary_var.set(
                f"{total} total  ·  {done} completed  ·  {pending} pending  ·  "
                f"{attention} need attention  ·  {skipped} skipped"
            )
        else:
            self.batch_summary_var.set("No batch loaded yet.")

    def _batch_process_selected(self):
        sel = self.batch_tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Click a row in the list first.")
            return
        index = int(sel[0])
        item = self.batch_queue[index]
        if item["status"] == "Completed":
            if not messagebox.askyesno(
                "Already completed",
                f"{item.get('employee_name') or item['employee_id']} was already processed and signed. "
                "Regenerate their PDF anyway?",
            ):
                return
        self._batch_auto_advance = False
        self._batch_load_row(index)

    def _batch_skip_selected(self):
        sel = self.batch_tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Click a row in the list first.")
            return
        index = int(sel[0])
        self.batch_queue[index]["status"] = "Skipped"
        if index == self._batch_active_index:
            self._batch_clear_panel()
        self._batch_refresh_tree()

    def _batch_requeue_selected(self):
        sel = self.batch_tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Click a row in the list first.")
            return
        index = int(sel[0])
        self.batch_queue[index]["status"] = "Pending"
        self._batch_refresh_tree()

    # ---------------------------------------------------------- auto-advance
    def _batch_process_all(self):
        if not self.batch_queue:
            messagebox.showinfo("No batch loaded", "Import an Excel file first.")
            return
        self._batch_auto_advance = True
        self.process_all_button.pack_forget()
        self.stop_auto_button.pack(side="right")
        self._batch_advance_auto()

    def _batch_stop_auto(self):
        self._batch_auto_advance = False
        self.stop_auto_button.pack_forget()
        self.process_all_button.pack(side="right")

    def _batch_advance_auto(self):
        next_index = next((i for i, it in enumerate(self.batch_queue) if it["status"] == "Pending"), None)
        if next_index is None:
            self._batch_auto_advance = False
            self.stop_auto_button.pack_forget()
            self.process_all_button.pack(side="right")
            done = sum(1 for it in self.batch_queue if it["status"] == "Completed")
            skipped = sum(1 for it in self.batch_queue if it["status"] == "Skipped")
            attention = sum(1 for it in self.batch_queue if it["status"] == "Needs Attention")
            if attention:
                messagebox.showinfo(
                    "Auto-advance paused",
                    f"{attention} person/people need attention (unrecognized Type or failed AD lookup) - "
                    "select them from the list above to fix and continue. Everyone else is done: "
                    f"{done} completed, {skipped} skipped.",
                )
            else:
                messagebox.showinfo("Batch complete", f"All done: {done} completed, {skipped} skipped.")
            return
        self._batch_load_row(next_index)

    # ---------------------------------------------------------- panel
    def _batch_clear_panel(self):
        self._batch_active_index = None
        for child in self.batch_person_frame.winfo_children():
            child.destroy()
        self.batch_person_frame.pack_forget()
        self.batch_placeholder.pack(anchor="w", pady=20)

    def _batch_load_row(self, index):
        self._batch_active_index = index
        item = self.batch_queue[index]
        if "_lookup_source" not in item:
            # Excel already supplied a name at import time -> "Import" is
            # the accurate lookup source unless/until SSRS or AD overrides
            # it below.
            item["_lookup_source"] = "Import" if item.get("employee_name") else ""

        self.batch_placeholder.pack_forget()
        for child in self.batch_person_frame.winfo_children():
            child.destroy()
        self.batch_person_frame.pack(fill="both", expand=True)

        self.b_signature_paths = {"employee_signature_path": None, "asset_receiver_signature_path": None}
        self.b_sig_preview_image = {}
        self.b_output_path_override = None

        parent = self.batch_person_frame
        ttk.Label(
            parent, text=f"Now Processing: {item.get('employee_name') or '(looking up...)'}   -   ID {item['employee_id']}",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", pady=(0, 4))
        ttk.Label(parent, text=f"Row {item['row_number']} in the imported file.", foreground="#666").pack(
            anchor="w", pady=(0, 14)
        )

        bu_match = self._batch_resolve_bu(item)
        bu_row = ttk.Frame(parent)
        bu_row.pack(anchor="w", pady=(0, 10))
        ttk.Label(bu_row, text="Business Unit:").pack(side="left")
        self.b_bu_value_label = ttk.Label(
            bu_row, text=(bu_match["name"] if bu_match else "(none - pick one at the top of this tab)"),
            font=("Segoe UI", 10, "bold"),
        )
        self.b_bu_value_label.pack(side="left", padx=6)

        ttk.Separator(parent).pack(fill="x", pady=10)

        details_box = ttk.LabelFrame(parent, text="Employee Details", padding=10)
        details_box.pack(fill="x", pady=(0, 12))
        self.b_emp_name_var = tk.StringVar(value=item.get("employee_name") or "")
        self.b_manager_name_var = tk.StringVar(value=item.get("manager_name") or "")
        ttk.Label(details_box, text="Employee Name:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(details_box, textvariable=self.b_emp_name_var, width=40).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(details_box, text="Manager Name:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(details_box, textvariable=self.b_manager_name_var, width=40).grid(row=1, column=1, sticky="w", padx=6)
        self.b_identity_hint = ttk.Label(details_box, text="", foreground="#a05a00", wraplength=620)
        self.b_identity_hint.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        # Same automatic SSRS check as the Single Person tab, near the
        # identity fields here since bulk rows don't have their own manual
        # "Lookup (AD)" button - AD lookup already runs automatically for a
        # row with no name yet (see below); this just tries the faster/
        # no-network SSRS cache first, before falling back to it.
        self.b_ssrs_check_label = ttk.Label(details_box, text="", foreground="#888")
        self.b_ssrs_check_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=(2, 0))

        contact_row = ttk.Frame(parent)
        contact_row.pack(anchor="w", pady=(0, 10))
        ttk.Label(contact_row, text="Contact Number:").pack(side="left")
        self.b_contact_var = tk.StringVar(value=item.get("contact_number") or "")
        ttk.Entry(contact_row, textvariable=self.b_contact_var, width=24).pack(side="left", padx=8)

        ttk.Separator(parent).pack(fill="x", pady=10)

        ttk.Label(parent, text="Submission Type:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        all_types = self.config_data.get("submission_type_row1", []) + self.config_data.get("submission_type_row2", [])
        self.b_submission_type_var = tk.StringVar(value=item.get("type") if item.get("type_recognized") else "")
        type_box = ttk.Frame(parent)
        type_box.pack(anchor="w", pady=(4, 0))
        for i, label in enumerate(all_types):
            ttk.Radiobutton(
                type_box, text=label, value=label, variable=self.b_submission_type_var,
                command=self._batch_on_submission_type_changed,
            ).grid(row=i // 3, column=i % 3, sticky="w", padx=10, pady=4)
        self.b_type_hint = ttk.Label(
            parent,
            text="" if item.get("type_recognized")
            else f"Type '{item.get('type')}' from the file wasn't recognized - please pick one above.",
            foreground="#a05a00",
        )
        self.b_type_hint.pack(anchor="w", pady=(4, 10))

        self.b_asset_details_frame = ttk.Frame(parent)
        self.b_asset_details_frame.pack(fill="x", anchor="w")
        self._batch_rebuild_asset_details()

        ttk.Separator(parent).pack(fill="x", pady=14)

        sig_row = ttk.Frame(parent)
        sig_row.pack(fill="x", anchor="w")
        emp_sig = ttk.Frame(sig_row)
        emp_sig.pack(side="left", padx=(0, 40), anchor="n")
        self._batch_build_signature_widget(emp_sig, "employee_signature_path", "Employee Signature", "Employee: sign here")
        ar_sig = ttk.Frame(sig_row)
        ar_sig.pack(side="left", anchor="n")
        self._batch_build_signature_widget(
            ar_sig, "asset_receiver_signature_path", "Asset Receiver Signature", "Asset Receiver: sign here"
        )

        ttk.Separator(parent).pack(fill="x", pady=14)

        # No manual Save-to/path/Choose Location here either - same
        # automatic root-folder + subfolder + filename as the Single
        # Person tab (see _batch_compute_default_output_path).
        self.b_output_path_override = None

        def _batch_refresh_save_path(*_args):
            pass  # no-op; kept callable since other code still calls it
        self._batch_refresh_save_path = _batch_refresh_save_path

        action_row = ttk.Frame(parent)
        action_row.pack(anchor="w", pady=(4, 20))
        ttk.Button(action_row, text="Save & Mark Complete", command=self._batch_save_and_complete).pack(side="left")
        ttk.Button(action_row, text="Skip This Person", command=self._batch_skip_active).pack(side="left", padx=10)
        ttk.Button(action_row, text="Cancel", command=self._batch_cancel_active).pack(side="left")

        self._batch_refresh_tree()

        if not item.get("employee_name"):
            # SSRS check first (Enhancement 7 relocation, no network call -
            # checked against the already-downloaded/cached report): only
            # fall through to the slower AD Lookup if SSRS has no match.
            if self._batch_try_ssrs_autofill_identity(index):
                self._batch_check_identity_resolved()
            else:
                self._batch_run_lookup_for_active()
        else:
            self._batch_check_identity_resolved()

    def _batch_try_ssrs_autofill_identity(self, index):
        """Returns True (and fills the name/asset fields immediately) if the
        SSRS asset report cache has a row for this employee_id; False if
        not, so the caller falls back to AD Lookup exactly as before."""
        item = self.batch_queue[index]
        emp_id = (item.get("employee_id") or "").strip()
        ssrs_idx = self.ssrs_state.get("index", {})
        record = ssrs_idx.get(emp_id) or ssrs_idx.get("".join(c for c in emp_id if c.isdigit()))
        if not hasattr(self, "b_ssrs_check_label"):
            return False
        if not record:
            self._b_ssrs_match_status = "Not Found" if self.ssrs_state.get("status") == SSRS_STATUS_AVAILABLE else ""
            if self.ssrs_state.get("status") == SSRS_STATUS_AVAILABLE:
                self.b_ssrs_check_label.config(text="SSRS: no match - checking AD...", foreground="#888")
            else:
                self.b_ssrs_check_label.config(text="")
            return False
        self._b_ssrs_match_status = "Matched"
        item["_ssrs_record"] = record
        item["_lookup_source"] = "SSRS"
        if record.get("employee_name"):
            self.b_emp_name_var.set(record["employee_name"])
        self.b_ssrs_check_label.config(
            text=f"SSRS: matched - {record.get('employee_name', '') or emp_id}", foreground="#1a7f37",
        )
        self.status_var.set(f"Batch: {self.b_emp_name_var.get() or emp_id} found via SSRS.")
        self._apply_ssrs_record_to_batch_asset_fields(record)
        # Same manager-name AD fallback as the Single Person tab - SSRS has
        # no manager column, so if the imported Excel data didn't supply
        # one either, quietly fetch it from AD in the background.
        self._batch_fetch_manager_via_ad_fallback(index, emp_id)
        return True

    def _batch_fetch_manager_via_ad_fallback(self, index, emp_id):
        if self.b_manager_name_var.get().strip():
            return
        uses_webview = _uses_webview_lookup(self.config_data)

        def worker():
            try:
                if uses_webview:
                    candidates = self.browser_session.search(emp_id)
                else:
                    candidates = [fetch_employee_details(emp_id, self.config_data)]
                self.after(0, self._batch_on_manager_fallback_result, index, candidates)
            except Exception:
                logger.debug("Batch manager AD fallback: lookup failed for %r (non-fatal)", emp_id, exc_info=True)

        threading.Thread(target=worker, daemon=True).start()

    def _batch_on_manager_fallback_result(self, index, candidates):
        if index != self._batch_active_index or not candidates:
            return
        if self.b_manager_name_var.get().strip():
            return
        manager = candidates[0].get("manager_name", "")
        if manager:
            self.b_manager_name_var.set(manager)

    def _apply_ssrs_record_to_batch_asset_fields(self, record):
        """Bulk-tab equivalent of _apply_ssrs_record_to_asset_fields - only
        meaningful for New Hire-type submissions, since that's the only
        case SSRS has issued-asset data for. Safe to call before a
        submission type/asset-details widgets exist yet (hasattr-guarded)."""
        sub_type = self.b_submission_type_var.get() if hasattr(self, "b_submission_type_var") else ""
        if sub_type not in self.config_data.get("ssrs_asset_report", {}).get("new_hire_submission_types", []):
            return
        if record.get("serial_number") and hasattr(self, "b_new_serial_var"):
            self.b_new_serial_var.set(record["serial_number"])
        summary = self._format_ssrs_record(record)
        if summary and hasattr(self, "b_assets_other_var"):
            existing = self.b_assets_other_var.get().strip()
            self.b_assets_other_var.set(f"{existing}; {summary}" if existing else summary)
        if hasattr(self, "b_ssrs_autofill_hint"):
            self.b_ssrs_autofill_hint.config(text="Filled from SSRS.", foreground="#1a7f37")

    def _batch_on_submission_type_changed(self):
        self._batch_rebuild_asset_details()
        if hasattr(self, "_batch_refresh_save_path"):
            self._batch_refresh_save_path()

    def _batch_rebuild_asset_details(self):
        for child in self.b_asset_details_frame.winfo_children():
            child.destroy()
        parent = self.b_asset_details_frame

        sub_type = self.b_submission_type_var.get()

        known_serial = ""
        active_emp_id = ""
        if self._batch_active_index is not None:
            active_item = self.batch_queue[self._batch_active_index]
            known_serial = active_item.get("laptop_serial_number") or ""
            active_emp_id = active_item.get("employee_id") or ""

        ttk.Label(parent, text="Asset Details", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(6, 2))

        # A New Hire has no existing device to return, so an imported
        # laptop_serial_number for that type is the device being ISSUED -
        # it prefills "New Device Serial Number" instead of "Current",
        # which is hidden entirely for New Hire (see show_current_serial
        # below and no_current_asset_submission_types in CONFIG).
        no_current_types = self.config_data.get("no_current_asset_submission_types", [])
        show_current_serial = sub_type not in no_current_types
        self.b_current_serial_var = tk.StringVar(value=known_serial if show_current_serial else "")
        self.b_new_serial_var = tk.StringVar(value=known_serial if not show_current_serial else "")
        self.b_assets_other_var = tk.StringVar(value="")
        self.b_asset_pending_var = tk.StringVar(value="")
        self.b_asset_checkbox_vars = {}

        if not sub_type:
            ttk.Label(parent, text="(select a submission type above)", foreground="#888").pack(anchor="w")
            return

        prefill_note = " (from the imported file - check it, then edit if needed)" if known_serial else ""

        if show_current_serial:
            row1 = ttk.Frame(parent)
            row1.pack(anchor="w", pady=4)
            ttk.Label(row1, text="Current Device Serial Number:").pack(side="left")
            ttk.Entry(row1, textvariable=self.b_current_serial_var, width=30).pack(side="left", padx=8)
            if sub_type in self.config_data.get("tracit_submission_types", []):
                ttk.Button(
                    row1, text="Open TracIT",
                    command=lambda: self._open_tracit(active_emp_id, self.b_current_serial_var.get().strip()),
                ).pack(side="left", padx=(4, 0))
            if prefill_note:
                ttk.Label(row1, text=prefill_note, foreground="#888").pack(side="left")

        # Enhancement 7 (relocated) - same automatic SSRS check as the
        # single-tab form; matching is done as soon as the row loads (see
        # _batch_try_ssrs_autofill_identity), this just re-displays/
        # re-applies that cached result whenever these fields are rebuilt
        # (e.g. after a submission type change).
        if sub_type in self.config_data.get("ssrs_asset_report", {}).get("new_hire_submission_types", []):
            ssrs_row = ttk.Frame(parent)
            ssrs_row.pack(anchor="w", pady=(0, 4))
            ttk.Label(ssrs_row, text="SSRS Report:", foreground="#666").pack(side="left")
            self.b_ssrs_autofill_hint = ttk.Label(ssrs_row, text="", foreground="#888")
            self.b_ssrs_autofill_hint.pack(side="left", padx=(6, 0))
            cached_record = None
            if self._batch_active_index is not None:
                cached_record = self.batch_queue[self._batch_active_index].get("_ssrs_record")
            if cached_record:
                self._apply_ssrs_record_to_batch_asset_fields(cached_record)

        row2 = ttk.Frame(parent)
        row2.pack(anchor="w", pady=4)
        ttk.Label(row2, text="New Device Serial Number:").pack(side="left")
        ttk.Entry(row2, textvariable=self.b_new_serial_var, width=30).pack(side="left", padx=8)

        ttk.Label(parent, text="List of Assets Issued/Return:").pack(anchor="w", pady=(14, 2))
        checklist_frame = ttk.Frame(parent)
        checklist_frame.pack(anchor="w")
        items = self.config_data.get("asset_checklist_items", [])
        for i, item_name in enumerate(items):
            var = tk.BooleanVar()
            self.b_asset_checkbox_vars[item_name] = var
            ttk.Checkbutton(checklist_frame, text=item_name, variable=var).grid(
                row=i // 3, column=i % 3, sticky="w", padx=8, pady=3
            )

        row3 = ttk.Frame(parent)
        row3.pack(anchor="w", pady=(10, 4))
        ttk.Label(row3, text="Others (specify):").pack(side="left")
        ttk.Entry(row3, textvariable=self.b_assets_other_var, width=40).pack(side="left", padx=8)

        row4 = ttk.Frame(parent)
        row4.pack(anchor="w", pady=4)
        ttk.Label(row4, text="Asset Pending for Submission (if any):").pack(side="left")
        ttk.Entry(row4, textvariable=self.b_asset_pending_var, width=40).pack(side="left", padx=8)

    def _batch_build_signature_widget(self, parent, data_key, title, capture_title):
        ttk.Label(parent, text=title, font=("Segoe UI", 11, "bold")).pack(anchor="w")
        preview_label = ttk.Label(parent, text="(not signed yet)", relief="groove", width=40, anchor="center")
        preview_label.pack(pady=(8, 8))

        def show_preview(path):
            try:
                img = Image.open(path)
                img.load()
                # See the matching comment in _build_signature_section -
                # normalize away from 1-bit/palette modes a Topaz pad's
                # bitmap can come in as, so the thumbnail actually
                # renders instead of showing up blank with no error.
                if img.mode not in ("RGB", "RGBA"):
                    img = img.convert("RGB")
                img.thumbnail((280, 110))
                photo = ImageTk.PhotoImage(img)
                self.b_sig_preview_image[data_key] = photo
                preview_label.image = photo  # belt-and-suspenders against GC
                preview_label.config(image=photo, text="", compound="image")
            except Exception:
                logger.exception("Signature preview (bulk): could not render thumbnail for %s (%s)", data_key, path)
                preview_label.config(text="(signature captured)")

        def sign_now():
            path = self.signature_service.capture(self, title=capture_title)
            if path:
                self.b_signature_paths[data_key] = path
                show_preview(path)
                self.status_var.set(f"{title} captured.")

        ttk.Button(parent, text="Sign Now...", command=sign_now).pack()

    # ---------------------------------------------------------- identity
    def _batch_run_lookup_for_active(self):
        index = self._batch_active_index
        item = self.batch_queue[index]
        self.status_var.set(f"Looking up {item['employee_id']}...")
        self.b_identity_hint.config(text="Looking up via AD...")
        uses_webview = _uses_webview_lookup(self.config_data)

        def worker():
            try:
                if uses_webview:
                    candidates = self.browser_session.search(item["employee_id"])
                else:
                    candidates = [fetch_employee_details(item["employee_id"], self.config_data)]
                self.after(0, self._batch_on_lookup_success, index, candidates)
            except Exception as exc:
                self.after(0, self._batch_on_lookup_failure, index, exc)

        threading.Thread(target=worker, daemon=True).start()

    def _batch_on_lookup_success(self, index, candidates):
        if index != self._batch_active_index:
            return  # operator already moved to a different person
        if not candidates:
            self._batch_mark_needs_attention("No AD match found - enter details manually.")
            return
        if len(candidates) == 1:
            chosen = candidates[0]
        else:
            chosen = self._prompt_candidate_selection(candidates)
            if chosen is None:
                self._batch_mark_needs_attention("Selection cancelled - enter details manually.")
                return
        self.b_emp_name_var.set(chosen.get("emp_name", ""))
        self.b_manager_name_var.set(chosen.get("manager_name", ""))
        self.batch_queue[index]["_lookup_source"] = "AD"
        self.b_identity_hint.config(text="")
        self.status_var.set(f"Batch: {self.b_emp_name_var.get()} found.")
        self._batch_check_identity_resolved()

    def _batch_on_lookup_failure(self, index, exc):
        logger.error(
            "Batch AD lookup failed (row index=%s, showing friendly message to operator): %s", index, exc,
            exc_info=(type(exc), exc, exc.__traceback__),
        )
        if index != self._batch_active_index:
            return
        self._batch_mark_needs_attention("⚠ Employee lookup unavailable. Please enter details manually.")

    def _batch_mark_needs_attention(self, message):
        index = self._batch_active_index
        self.batch_queue[index]["status"] = "Needs Attention"
        self.batch_queue[index]["_lookup_source"] = "Manual"
        self.b_identity_hint.config(text=message)
        self.status_var.set(f"Batch: {message}")
        self._batch_refresh_tree()

    def _batch_check_identity_resolved(self):
        index = self._batch_active_index
        item = self.batch_queue[index]
        if not self.b_submission_type_var.get():
            item["status"] = "Needs Attention"
            self.b_type_hint.config(
                text=f"Type '{item.get('type')}' from the file wasn't recognized - please pick one above."
            )
        elif item["status"] not in ("Completed", "Skipped"):
            item["status"] = "Pending"
        self._batch_refresh_tree()

    def _batch_on_default_bu_changed(self, event=None):
        """Keep the active person panel's Business Unit label in sync when the
        operator changes the batch-wide default BU after a row is already
        loaded. (The actual value used at save time is always resolved fresh
        via _batch_resolve_bu regardless, so this is a display-only fix.)"""
        if self._batch_active_index is None or not hasattr(self, "b_bu_value_label"):
            return
        item = self.batch_queue[self._batch_active_index]
        bu_match = self._batch_resolve_bu(item)
        self.b_bu_value_label.config(
            text=(bu_match["name"] if bu_match else "(none - pick one at the top of this tab)")
        )
        if hasattr(self, "_batch_refresh_save_path"):
            self._batch_refresh_save_path()

    # ---------------------------------------------------------- save/skip
    def _batch_resolve_bu(self, item):
        if item.get("business_unit"):
            match = next((t for t in self._bu_templates_cache if t["name"] == item["business_unit"]), None)
            if match:
                return match
        if self.bulk_bu_var.get():
            return next((t for t in self._bu_templates_cache if t["name"] == self.bulk_bu_var.get()), None)
        return None

    def _batch_compute_default_output_path(self, item):
        # Same Enhancement 5/6 structure and filename convention as the
        # single-tab version above - kept in sync deliberately.
        root = (self.user_settings.get("root_save_folder") or "").strip() or _resolve_path(
            self.config_data, "pdf_save_folder", "generated_forms"
        )
        emp_id = item["employee_id"]
        sub_type = self.b_submission_type_var.get().strip() if hasattr(self, "b_submission_type_var") else (item.get("type") or "")
        operator = self.user_settings.get("operator_name") or "Unspecified_Operator"
        today_str = date.today().strftime("%d-%b-%Y")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        compact_type = SUBMISSION_TYPE_TAGS.get(sub_type) or _safe_folder_name(sub_type)
        filename = f"{emp_id}_{compact_type}_{timestamp}.pdf"
        return os.path.join(
            root, _safe_folder_name(sub_type), today_str, _safe_folder_name(operator), filename
        )

    def _batch_validate_active(self):
        item = self.batch_queue[self._batch_active_index]
        emp_id = (item.get("employee_id") or "").strip()
        if not emp_id.isdigit():
            messagebox.showwarning(
                "Invalid Employee ID",
                f"Row {item.get('row_number')}'s Employee ID ({emp_id!r}) contains something other than "
                "digits - fix it in the imported Excel file and re-import, or Skip this person.",
            )
            return None
        bu_match = self._batch_resolve_bu(item)
        if not bu_match:
            messagebox.showwarning(
                "Business Unit required",
                "Select a Business Unit for this batch (top of this tab), or set one in the Excel "
                "file's 'Business Unit' column for this row.",
            )
            return None
        if not self.b_emp_name_var.get().strip():
            messagebox.showwarning("Missing Employee Name", "Enter the employee's name (AD lookup found nothing).")
            return None
        if not self.b_contact_var.get().strip():
            messagebox.showwarning("Missing contact number", "Enter a contact number.")
            return None
        if not self.b_submission_type_var.get():
            messagebox.showwarning("Missing selection", "Select a submission type.")
            return None
        if not self.b_signature_paths.get("employee_signature_path"):
            messagebox.showwarning("Signature required", "Capture the Employee Signature.")
            return None
        if not self.b_signature_paths.get("asset_receiver_signature_path"):
            messagebox.showwarning("Signature required", "Capture the Asset Receiver Signature.")
            return None
        return bu_match

    def _batch_save_and_complete(self):
        index = self._batch_active_index
        item = self.batch_queue[index]
        logger.info(
            "Batch generate clicked: index=%d emp_id=%r bu=%r",
            index, item.get("employee_id"), getattr(self, "bulk_bu_var", None) and self.bulk_bu_var.get(),
        )
        bu_match = self._batch_validate_active()
        if not bu_match:
            logger.info("Batch generate: validation failed for index=%d, stopping", index)
            return

        today_str = date.today().strftime("%d-%b-%Y")
        emp_id = item["employee_id"]
        emp_name = self.b_emp_name_var.get().strip()
        manager_name = self.b_manager_name_var.get().strip()
        contact_number = self.b_contact_var.get().strip()
        chosen_type = self.b_submission_type_var.get()

        field_values = {
            "Date": today_str, "EmpID": emp_id, "EmpName": emp_name,
            "ContactNumber": contact_number, "ManagerName": manager_name,
            "LastWorkingDate": item.get("last_working_date") or "",
        }

        checked_tags = {tag for label, tag in SUBMISSION_TYPE_TAGS.items() if label == chosen_type}
        current_serial = self.b_current_serial_var.get().strip()
        new_serial = self.b_new_serial_var.get().strip()
        assets_other = self.b_assets_other_var.get().strip()
        asset_pending = self.b_asset_pending_var.get().strip()
        assets_issued = [k for k, v in self.b_asset_checkbox_vars.items() if v.get()]

        field_values["CurrentSerial"] = current_serial
        field_values["NewSerial"] = new_serial
        field_values["Others"] = assets_other
        field_values["AssetPending"] = asset_pending
        checked_tags |= {tag for label, tag in ASSET_CHECKLIST_TAGS.items() if label in assets_issued}

        signature_images = {
            "EmployeeSignature": self.b_signature_paths.get("employee_signature_path"),
            "AssetReceiverSignature": self.b_signature_paths.get("asset_receiver_signature_path"),
        }

        output_path = self.b_output_path_override or self._batch_compute_default_output_path(item)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        operator_name = self.user_settings.get("operator_name") or ""
        operator_location = self.user_settings.get("location") or ""
        ssrs_match_status = getattr(self, "_b_ssrs_match_status", "") or ""

        tmp_dir = tempfile.mkdtemp(prefix="itasset_")
        filled_docx = os.path.join(tmp_dir, "filled.docx")
        unsigned_pdf = os.path.join(tmp_dir, "unsigned.pdf")

        self.status_var.set(f"Filling document for {emp_name}...")
        self.update_idletasks()
        try:
            fill_docx(bu_match["fillable_path"], filled_docx, field_values, checked_tags, signature_images)
            self.status_var.set("Exporting to PDF via Word...")
            self.update_idletasks()
            convert_docx_to_pdf_via_word(filled_docx, unsigned_pdf)

            # Enhancement 13/18 - visible "Signed On: <timestamp>" caption
            # under both signatures, before signing (see the single-tab
            # _on_generate for why the ordering matters).
            _stamp_signed_on_captions(unsigned_pdf, _find_signature_field_boxes(unsigned_pdf))

            # Enhancement 19 - PDF metadata, before signing (see the
            # single-tab _on_generate for why the ordering matters).
            _set_pdf_metadata(unsigned_pdf, {
                "Title": f"IT Asset Submission Acknowledgement - {emp_name}",
                "Author": operator_name or "IT Asset Submission Acknowledgement System",
                "Subject": chosen_type,
                "EmployeeName": emp_name, "EmployeeID": emp_id, "ManagerName": manager_name,
                "SubmissionType": chosen_type, "GeneratedBy": operator_name,
                "GeneratedDate": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ProbTicket": "", "SCID": "",
            })

            self.status_var.set("Applying cryptographic signatures...")
            self.update_idletasks()
            # Enhancement 12 - Adobe's own Signature Properties dialog has
            # no dedicated Employee ID field, so it's folded into the
            # Reason text (along with the employee name) for both
            # signatures, alongside the real Location and the automatic
            # signing timestamp pyHanko already embeds.
            base_reason = self.config_data.get("signing_reason") or "IT Asset Acknowledgement"
            emp_reason = f"{base_reason} - {emp_name} (Employee ID: {emp_id})"
            sign_pdf_with_signatures(
                unsigned_pdf, output_path,
                [
                    {"field_name": "EmployeeSignature", "display_name": emp_name,
                     "location": operator_location, "reason": emp_reason},
                    {"field_name": "AssetReceiverSignature",
                     "display_name": self.config_data.get("asset_receiver_display_name", "Asset Receiver"),
                     "location": operator_location, "reason": emp_reason},
                ],
                self.config_data,
            )
        except (FillDocxError, WordConversionError, SigningError) as exc:
            logger.error("Batch generate failed for emp_id=%r: %s", emp_id, exc)
            messagebox.showerror("Generation Error", str(exc))
            self.status_var.set("Generation failed.")
            _write_audit_log_entry(
                operator=operator_name, employee_id=emp_id, submission_type=chosen_type,
                ssrs_match_status=ssrs_match_status, lookup_source=item.get("_lookup_source", ""),
                pdf_generated=False, pdf_location="", email_draft_created=False,
            )
            return
        except Exception as exc:
            logger.exception("Batch generate failed unexpectedly for emp_id=%r", emp_id)
            messagebox.showerror("Generation Error", f"Unexpected error: {exc}")
            self.status_var.set("Generation failed.")
            _write_audit_log_entry(
                operator=operator_name, employee_id=emp_id, submission_type=chosen_type,
                ssrs_match_status=ssrs_match_status, lookup_source=item.get("_lookup_source", ""),
                pdf_generated=False, pdf_location="", email_draft_created=False,
            )
            return
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

        item["status"] = "Completed"
        item["output_path"] = output_path
        item["employee_name"] = emp_name
        item["manager_name"] = manager_name
        item["type"] = chosen_type
        item["type_recognized"] = True

        email_draft_created = False
        if _is_email_draft_type(self.config_data, chosen_type):
            email_data = {
                "emp_id": emp_id, "emp_name": emp_name, "manager_name": manager_name,
                "submission_type": chosen_type, "date": today_str,
                "current_device_serial": current_serial, "assets_issued": assets_issued,
                "assets_other": assets_other, "asset_pending": asset_pending,
                "email_action_label": self.config_data.get("email_templates", {}).get(chosen_type, chosen_type),
            }
            try:
                open_draft(email_data, to_addresses=self.config_data.get("lwd_email_to", ""), attachment_path=output_path)
                email_draft_created = True
            except RuntimeError as exc:
                messagebox.showwarning("Email draft not opened", str(exc))

        _write_audit_log_entry(
            operator=operator_name, employee_id=emp_id, submission_type=chosen_type,
            ssrs_match_status=ssrs_match_status, lookup_source=item.get("_lookup_source", ""),
            pdf_generated=True, pdf_location=output_path, email_draft_created=email_draft_created,
        )
        if hasattr(self, "_history_refresh"):
            self._history_refresh()
        if hasattr(self, "_recent_docs_refresh"):
            self._recent_docs_refresh()

        self.status_var.set(f"Completed: {emp_name} -> {output_path}")
        self._batch_clear_panel()
        self._batch_refresh_tree()

        if self._batch_auto_advance:
            self._batch_advance_auto()

    def _batch_skip_active(self):
        index = self._batch_active_index
        self.batch_queue[index]["status"] = "Skipped"
        self._batch_clear_panel()
        self._batch_refresh_tree()
        if self._batch_auto_advance:
            self._batch_advance_auto()

    def _batch_cancel_active(self):
        self._batch_auto_advance = False
        self.stop_auto_button.pack_forget()
        self.process_all_button.pack(side="right")
        self._batch_clear_panel()
        self._batch_refresh_tree()

    # ------------------------------------------------------------ close
    def _on_close(self):
        logger.info("WizardApp: closing (user closed the window)")
        try:
            self.browser_session.shutdown()
        except Exception:
            logger.debug("WizardApp: browser_session.shutdown() failed (non-fatal)", exc_info=True)
        self.destroy()
        logger.info("WizardApp: closed")



# ============================================================================
# ENTRY POINT
# ============================================================================
def _set_windows_dpi_awareness():
    """Without this, Tkinter is NOT DPI-aware on Windows, so Windows
    applies its own bitmap-stretching compatibility scaling to the whole
    window on any scaled display (125%/150%/etc. - extremely common on
    laptops). A window created at a nominal "950x860" then actually
    renders much taller in real screen pixels than that number implies,
    which was pushing the Save/Generate button bar below the visible
    screen area entirely - looking exactly like "the Save button isn't
    there" even though it's in the layout the whole time. Declaring
    DPI-awareness up front (must happen before any Tk window is created)
    makes Windows hand Tkinter real pixels 1:1 instead of scaling it,
    which is the standard fix for this class of bug. Safe to skip
    silently on non-Windows or older Windows without these APIs."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()  # older Windows fallback
    except Exception:
        logger.debug("Could not set Windows DPI awareness (non-fatal)", exc_info=True)


if __name__ == "__main__":
    _log_path = _setup_logging()
    _enable_os_trust_store()
    if len(sys.argv) > 1 and sys.argv[1] == "--webview-helper":
        # Relaunched as the embedded-browser AD lookup helper - see
        # BrowserLookupSession._build_command() above for how/why.
        _run_webview_helper()
    else:
        _set_windows_dpi_awareness()
        app = WizardApp()
        app.mainloop()
