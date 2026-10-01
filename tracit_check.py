=== TracIT EUC report - export details (secrets masked) ===
Collected: 2026-10-01 13:08   Python 3.14.7
Looked in: C:\Users\srajaamu\Desktop, C:\Users\srajaamu\Downloads, C:\Users\srajaamu\OneDrive - UHG\Desktop

[1] EXPORT REQUEST  (from tracit.optum.com.har)
    22 requests in the log

  --- Candidate 1 (match score 2) ---
    Method      : GET
    URL         : https://tracit.optum.com/api/ham/api/ReportsLookUp/GetDeviceTypeByReportType?deviceCategory=2
    Status      : 200
    Content-Type: application/json; charset=utf-8
    Disposition : 
    Size (bytes): 602
    Request headers (secret values masked):
      Accept: application/json, text/plain, */*
      Accept-Encoding: gzip, deflate, br, zstd
      Accept-Language: en-US,en;q=0.9,en-IN;q=0.8
      Access-Control-Allow-Headers: Content-Type
      Access-Control-Allow-Methods: *
      Access-Control-Allow-Origin: *
      Connection: keep-alive
      DNT: 1
      Host: tracit.optum.com
      Referer: https://tracit.optum.com/actionable-insights/eucreport
      Sec-Fetch-Dest: empty
      Sec-Fetch-Mode: cors
      Sec-Fetch-Site: same-origin
      User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36 Edg/152.0.0.0
      sec-ch-ua: "Chromium";v="152", "Not?A_Brand";v="24", "Microsoft Edge";v="152"
      sec-ch-ua-mobile: ?0
      sec-ch-ua-platform: "Windows"

  --- Candidate 2 (match score 2) ---
    Method      : GET
    URL         : https://tracit.optum.com/api/ham/api/ReportsLookUp/GetLCSByReportType?deviceCategory=2
    Status      : 200
    Content-Type: application/json; charset=utf-8
    Disposition : 
    Size (bytes): 16011
    Request headers (secret values masked):
      Accept: application/json, text/plain, */*
      Accept-Encoding: gzip, deflate, br, zstd
      Accept-Language: en-US,en;q=0.9,en-IN;q=0.8
      Access-Control-Allow-Headers: Content-Type
      Access-Control-Allow-Methods: *
      Access-Control-Allow-Origin: *
      Connection: keep-alive
      DNT: 1
      Host: tracit.optum.com
      Referer: https://tracit.optum.com/actionable-insights/eucreport
      Sec-Fetch-Dest: empty
      Sec-Fetch-Mode: cors
      Sec-Fetch-Site: same-origin
      User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36 Edg/152.0.0.0
      sec-ch-ua: "Chromium";v="152", "Not?A_Brand";v="24", "Microsoft Edge";v="152"
      sec-ch-ua-mobile: ?0
      sec-ch-ua-platform: "Windows"

  --- Candidate 3 (match score 2) ---
    Method      : GET
    URL         : https://tracit.optum.com/api/ham/api/ReportsLookUp/GetAllSpecialProjectHandleTypes
    Status      : 200
    Content-Type: application/json; charset=utf-8
    Disposition : 
    Size (bytes): 5026
    Request headers (secret values masked):
      Accept: application/json, text/plain, */*
      Accept-Encoding: gzip, deflate, br, zstd
      Accept-Language: en-US,en;q=0.9,en-IN;q=0.8
      Access-Control-Allow-Headers: Content-Type
      Access-Control-Allow-Methods: *
      Access-Control-Allow-Origin: *
      Connection: keep-alive
      DNT: 1
      Host: tracit.optum.com
      Referer: https://tracit.optum.com/actionable-insights/eucreport
      Sec-Fetch-Dest: empty
      Sec-Fetch-Mode: cors
      Sec-Fetch-Site: same-origin
      User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36 Edg/152.0.0.0
      sec-ch-ua: "Chromium";v="152", "Not?A_Brand";v="24", "Microsoft Edge";v="152"
      sec-ch-ua-mobile: ?0
      sec-ch-ua-platform: "Windows"

  --- Candidate 4 (match score 2) ---
    Method      : POST
    URL         : https://tracit.optum.com/api/ham/api/Reports/GetEUCReportExtractNew
    Status      : 200
    Content-Type: application/json; charset=utf-8
    Disposition : 
    Size (bytes): 771412
    Request headers (secret values masked):
      Accept: application/json, text/plain, */*
      Accept-Encoding: gzip, deflate, br, zstd
      Accept-Language: en-US,en;q=0.9,en-IN;q=0.8
      Access-Control-Allow-Headers: Content-Type
      Access-Control-Allow-Methods: *
      Access-Control-Allow-Origin: *
      Connection: keep-alive
      Content-Length: 1525
      Content-Type: application/json
      DNT: 1
      Host: tracit.optum.com
      Origin: https://tracit.optum.com
      Referer: https://tracit.optum.com/actionable-insights/eucreport
      Sec-Fetch-Dest: empty
      Sec-Fetch-Mode: cors
      Sec-Fetch-Site: same-origin
      User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36 Edg/152.0.0.0
      sec-ch-ua: "Chromium";v="152", "Not?A_Brand";v="24", "Microsoft Edge";v="152"
      sec-ch-ua-mobile: ?0
      sec-ch-ua-platform: "Windows"
    Payload (application/json):
      {"pageSize":6000,"pageNumber":1,"assetType":[],"lifecycleStatus":[],"lifecycleSubStatus":[],"assetLocationCountry":[],"specialProjectHandlingType":[],"assetScanAging":[],"assetAging":[],"companyName":[],"assignedLocation":["II033"],"assetLocation":[],"aud":[],"serialNumber":[],"assetTag":[],"machineName":[],"primaryUserEmployeeId":[],"primaryUserMSId":[],"assignedOwnerEmployeeId":[],"assignedOwnerUserId":[],"poNumber":[],"manufacturer":[],"audOwnerMSId":[],"modelNM":null,"partNumber":null,"osName":null,"computraceAgent":[],"computraceEventStatus":[],"lastLogonDateFrom":null,"lastLogonDateTo":null,"lifecycleStatusDateFrom":null,"lifecycleStatusDateTo":null,"columns":["serialNumber","assetTag","deviceType","hostName","partNumber","modelNm","manufacturerNm","osName","lifecycleStatus","lifecycleSubStatus","poNumber","lifecycleStatusUpdateDate","assignedOwnerEmpid","assignedOwnerName","assignedOwnerUserID","computerUser","computerUserEMPID","assetLocation","assetLocationCountry","assetLocationCity","assignedLocation","companyName","isIntegrated","uhgOwned","audOwner","audOwnerMsId","audRequestNumber","audEffectiveDate","audFollowupDate","audComment","comment","specialProjectHandlingType","chargebackType","assetAging","monthsInService","purchaseInvoiceDate","assetScanAging","isAUD","lastScanDate","lastScanSource","lastLogonDate","lastLogonUserId","computraceAgent","computraceEventStatus","computraceLastCall","audCategoryName","previousLifecycleStatus","previousLifecycleDate","eucDeviceId","svpVendorName"]}

    Requests just before the export (to spot a two-step export):
      GET https://tracit.optum.com/api/ham/api/Common/SavedViewV1/GetSavedView?projectId=1 -> 200 application/json; charset=utf-8

[2] DOES THE LINK WORK WITH ONLY YOUR WINDOWS LOGIN?  (no browser cookies - same as the app)
    Auth used : Windows login (requests-negotiate-sspi)
    HTTP 403, Content-Type: text/html, 179 bytes
    Result    : HTML page (probably a LOGIN page) - needs browser sign-in

[3] REPORT FILE (headers + row count only - no data values)
    File: EUCreport (8).xlsx  (117 KB, saved 2026-10-01 10:04)
    Header row is row 1; about 543 data rows
    Columns (with the SHAPE of the first data row - values are not copied):
       1. Serial Number                       code x10
       2. Asset Tag                           digits x7
       3. Device Type                         text x6
       4. Host Name                           code x15
       5. Part Number                         text x11
       6. Model Name                          text x15
       7. Manufacturer Name                   code x2
       8. OS Name                             text x21
       9. Lifecycle Status                    text x10
      10. Lifecycle Sub Status                (empty)
      11. PO Number                           code x9
      12. Lifecycle Status Update Date        date
      13. Assigned Owner Employee ID          (empty)
      14. Assigned Owner Name                 (empty)
      15. Assigned Owner User ID              (empty)
      16. Computer User                       (empty)
      17. Computer User Employee ID           (empty)
      18. Asset Location                      code x5
      19. Asset Location Country              code x3
      20. Asset Location City                 text x7
      21. Assigned Location                   code x5
      22. Company Name                        text x18
      23. Is Integrated                       text x4
      24. UHG Owned                           text x3
      25. AUD Owner                           (empty)
      26. AUD Owner MSID                      (empty)
      27. AUD Request Number                  (empty)
      28. AUD Effective Date                  (empty)
      29. AUD Followup Date                   (empty)
      30. AUD Comment                         (empty)
      31. Comment                             text x80
      32. Special Project Handling Type       (empty)
      33. Chargeback Type                     (empty)
      34. Asset Aging                         text x12
      35. Months In Service                   digits x2
      36. Purchase Invoice Date               date
      37. Asset Scan Aging                    text x15
      38. Is AUD                              text x7
      39. Last Scan Date                      date
      40. Last Scan Source                    text x7
      41. Last Logon Date                     date
      42. Last Logon User ID                  text x8
      43. Computrace Agent                    code x1
      44. Computrace Event Status             (empty)
      45. Computrace Last Call                date
      46. AUD Category                        (empty)
      47. Previous Lifecycle Status           text x6
      48. Previous Lifecycle Date             text x19
      49. EUC Device Id                       digits x7
      50. SVP Vendor Name                     code x5

=== end ===
