python check_servicenow.py --base-url https://yourcompany.service-now.com --table u_new_hire_onboarding --emp-id 900311850


"""
ServiceNow discovery/check script - run this ONCE on your Windows machine
(same one that runs the IT Asset Submission app) to find out the real
table name and field names for New Hire records, so I can wire the actual
ServiceNow module into app.py correctly instead of guessing.

WHAT IT DOES
    - Connects to ServiceNow using your current Windows login (the same
      SSPI/Windows-Integrated-Auth approach the app already uses for SSRS
      and AD - no password is ever typed in or stored).
    - Downloads a few sample records (or one specific Employee ID's record,
      if you give it one) with EVERY field ServiceNow returns.
    - Writes the full result to servicenow_dump.json next to this script,
      and also prints a short readable summary to the screen.
    - Does NOT write/change anything in ServiceNow - read-only (a GET
      request), completely safe to run.

HOW TO FIND YOUR TABLE NAME
    Open the New Hire list/record in ServiceNow in your browser, look at
    the URL. It usually looks like one of these:
        https://yourcompany.service-now.com/u_new_hire_onboarding_list.do
        https://yourcompany.service-now.com/nav_to.do?uri=/u_new_hire_onboarding_list.do
        https://yourcompany.service-now.com/now/nav/ui/classic/params/target/u_new_hire_onboarding_list.do
    The table name is the part right before "_list.do" (or "_form.do" if
    you're looking at one specific record) - in the example above, that's
        u_new_hire_onboarding
    If you're not sure, just run this script with --table left out of the
    command below the first time - it will explain what to try.

HOW TO RUN IT
    1. Make sure requests and requests-negotiate-sspi are installed
       (they're already in requirements.txt for the main app):
           pip install requests requests-negotiate-sspi

    2. Run it with your ServiceNow base URL and table name:
           python check_servicenow.py --base-url https://yourcompany.service-now.com --table u_new_hire_onboarding

       Optional: if you already know one real Employee ID that should be
       in that table right now, pass it too - the script will try several
       common field-name guesses to find the record for that specific
       person (much easier to read than 5 random rows):
           python check_servicenow.py --base-url https://yourcompany.service-now.com --table u_new_hire_onboarding --emp-id 900311850

    3. Send me the servicenow_dump.json file it creates (or paste its
       contents, or a screenshot of the printed summary) - that's all I
       need to fill in the real table/field names in app.py.

WHAT TO DO IF IT FAILS
    - "401 Unauthorized" / "403 Forbidden": your Windows account may not
      have API access to that table, or ServiceNow's REST API may need to
      be enabled for your account by your ServiceNow admin - ask them.
    - "Could not find table" / 404: the table name is wrong - double-check
      the URL as described above, or ask your ServiceNow admin for the
      exact Table API name (they may call it something like "New Hire
      Onboarding [u_new_hire_onboarding]" in Studio).
    - SSL certificate errors: same underlying issue as the SSRS one from
      before - run `pip install truststore` (already in requirements.txt)
      and try again.
"""
import argparse
import json
import sys

try:
    import requests
except ImportError:
    print("The 'requests' library is not installed. Run:\n    pip install requests")
    sys.exit(1)

try:
    from requests_negotiate_sspi import HttpNegotiateAuth
    HAS_SSPI = True
except ImportError:
    HAS_SSPI = False

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

# Common guesses for which ServiceNow field holds the Employee ID, tried in
# order when --emp-id is given. Whichever one actually finds a match tells
# us the real query_field to use in app.py's config.
EMPLOYEE_ID_FIELD_GUESSES = [
    "u_employee_id", "employee_id", "u_empl_id", "u_emp_id",
    "u_employee_number", "employee_number", "u_person_id", "number",
]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--base-url", required=True, help="e.g. https://yourcompany.service-now.com")
    parser.add_argument("--table", required=True, help="ServiceNow Table API name, e.g. u_new_hire_onboarding")
    parser.add_argument("--emp-id", default=None, help="A real Employee ID to look up (optional but recommended)")
    parser.add_argument("--limit", type=int, default=5, help="How many sample rows to dump if --emp-id isn't given (default 5)")
    parser.add_argument("--out", default="servicenow_dump.json", help="Output JSON file (default servicenow_dump.json)")
    args = parser.parse_args()

    if not HAS_SSPI:
        print("The 'requests_negotiate_sspi' library is not installed. Run:\n    pip install requests-negotiate-sspi")
        sys.exit(1)

    base_url = args.base_url.rstrip("/")
    table_url = f"{base_url}/api/now/table/{args.table}"
    auth = HttpNegotiateAuth()
    headers = {"Accept": "application/json"}

    def do_get(params, label):
        print(f"\n--- {label} ---")
        print(f"GET {table_url}")
        print(f"    params: {params}")
        try:
            resp = requests.get(table_url, params=params, auth=auth, headers=headers, timeout=20)
        except requests.exceptions.SSLError as exc:
            print(f"SSL certificate error: {exc}\n-> run: pip install truststore   (see the top of this script)")
            return None
        except requests.exceptions.ConnectionError as exc:
            print(f"Could not reach {base_url} - check VPN/network: {exc}")
            return None
        if resp.status_code == 401:
            print("401 Unauthorized - your Windows account isn't authenticating to ServiceNow's API. "
                  "Ask your ServiceNow admin whether REST API access is enabled for your account.")
            return None
        if resp.status_code == 403:
            print("403 Forbidden - authenticated, but no access to this table's API. Ask your ServiceNow admin.")
            return None
        if resp.status_code == 404:
            print(f"404 Not Found - '{args.table}' doesn't look like a real Table API name. "
                  "Double check the table name (see the how-to at the top of this script).")
            return None
        try:
            resp.raise_for_status()
        except Exception as exc:
            print(f"Request failed: {exc}\nResponse body: {resp.text[:2000]}")
            return None
        try:
            return resp.json().get("result", [])
        except Exception:
            print(f"Response wasn't valid JSON. First 2000 chars:\n{resp.text[:2000]}")
            return None

    rows = None
    matched_field = None

    if args.emp_id:
        for field in EMPLOYEE_ID_FIELD_GUESSES:
            rows = do_get({"sysparm_query": f"{field}={args.emp_id}", "sysparm_limit": "1",
                            "sysparm_display_value": "true"}, f"trying query_field={field!r}")
            if rows:
                matched_field = field
                print(f"\n*** FOUND A MATCH using field '{field}' - this is very likely your query_field. ***")
                break
        if not rows:
            print(
                f"\nNone of the common Employee ID field-name guesses matched Employee ID "
                f"{args.emp_id!r} in table '{args.table}'. Either that ID isn't in this table right "
                "now, or the real field has a different name than the ones this script tried "
                f"({', '.join(EMPLOYEE_ID_FIELD_GUESSES)}). Falling back to dumping "
                f"{args.limit} sample row(s) instead so you can see the real field names."
            )
            rows = do_get({"sysparm_limit": str(args.limit), "sysparm_display_value": "true"}, "sample rows (unfiltered)")
    else:
        rows = do_get({"sysparm_limit": str(args.limit), "sysparm_display_value": "true"}, "sample rows (unfiltered)")

    if not rows:
        print("\nNo rows returned - see any error above. Nothing written.")
        sys.exit(1)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump({"table": args.table, "query_field_that_matched": matched_field, "rows": rows}, f, indent=2)

    print(f"\nWrote {len(rows)} row(s) to {args.out} - please send me this file.")
    print("\n--- Quick summary of field names found (send this too if easier than the file) ---")
    all_fields = sorted({k for row in rows for k in row.keys()})
    for field in all_fields:
        sample_val = rows[0].get(field, "")
        if isinstance(sample_val, dict):
            sample_val = sample_val.get("display_value", sample_val)
        sample_str = str(sample_val)
        if len(sample_str) > 60:
            sample_str = sample_str[:57] + "..."
        print(f"  {field:35s} example: {sample_str}")


if __name__ == "__main__":
    main()
