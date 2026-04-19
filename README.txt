============================================================
 TELCO THREAT INTEL REPORT GENERATOR
 Windows Desktop Application
============================================================

QUICK START
-----------
1. Double-click  Generate_Report.bat  to launch
   (installs dependencies automatically on first run)

   OR

   Open Command Prompt in this folder and run:
     pip install -r requirements.txt
     python generate_report_gui.py

REQUIREMENTS
------------
- Windows 10 / 11
- Python 3.8 or higher  →  https://python.org/downloads
  (check "Add Python to PATH" during install)
- reportlab    (auto-installed by .bat)
- pyyaml       (auto-installed by .bat)
- python-docx  (auto-installed by .bat)

  If Python is not installed on Windows, the launcher will
  automatically try WSL (Windows Subsystem for Linux) as a
  fallback — no extra setup needed if WSL is available.

HOW TO USE
----------
The app has 6 tabs. Fill in what you have, leave blank what you don't.

TAB 1 — Report Info
    Title, subtitle, period, author, org, TLP level, threat level
    Output PDF Path  — where to save the PDF
    Output DOCX Path — where to save the Word document

TAB 2 — Executive Summary
    - Opening paragraph (overall assessment)
    - Top risks (one per line)
    - Key findings table (pipe-separated):
      Finding | Actor/Vector | Severity

TAB 3 — Threat Landscape
    Five domain sections: Signaling, 5G Core, Enterprise IT, Fraud, Supply Chain
    Each has: Overview | Incidents | Mitigations

    Incidents format (pipe-separated):
      Title | Date | Severity | Description
      Example: SS7 OTP Interception | Mar 2026 | CRITICAL | Active campaign targeting Indian subs

TAB 4 — Actors & CVEs
    Threat Actors (pipe-separated):
      Name | Origin | Target | Initial Access | TTPs | MITRE IDs
      Example: Salt Typhoon | China (PRC) | Backbone routers | Exploit IOS-XE | LotL | T1190,T1133

    Vulnerabilities (pipe-separated):
      CVE ID | Vendor/Product | Description | CVSS | Patch Priority
      Patch Priority: IMMEDIATE | HIGH | MEDIUM | LOW

TAB 5 — Malware & Breaches
    Malware (pipe-separated):
      Name | Date | Severity | Description | Recommendation
      Severity: CRITICAL | HIGH | MEDIUM

    Breaches (pipe-separated):
      Organization | Date | Sector | Records Exposed | Vector | Status
      Status: Confirmed | Under Investigation | Alleged | Contained

TAB 6 — IOCs
    IMPORTANT: Always defang IOCs!
      Replace . with [.] in IPs:      185.220.101[.]47
      Replace . with [.] in domains:  malware-site[.]com

    Format (pipe-separated):
      Type | Value | Campaign/Malware | Confidence | Action
      Action: BLOCK | ALERT | MONITOR | BLOCK at STP
      Confidence: HIGH | MEDIUM | LOW

    IOC Types: IPv4 | IPv6 | Domain | URL | File Hash (SHA256) |
               File Hash (MD5) | SCCP GT | MSISDN Range | User-Agent | Email

BUTTONS (bottom bar)
--------------------
  Save YAML Manifest  →  Save all form fields to a .yaml file for reuse
  Load YAML Manifest  →  Load a previously saved form
  GENERATE DOCX       →  Build and save a Word document (.docx)
  GENERATE PDF        →  Build and save the PDF report

OUTPUT FORMATS
--------------
PDF:
  Dark-themed professional report with colour-coded severity, tables,
  and TLP:AMBER header/footer. Default: ~/Downloads/Telco_TI_Report.pdf

DOCX (Word):
  Fully editable Word document with the same sections and tables.
  Clean professional formatting — navy/teal headings, colour-coded
  severity and patch priority columns. Use this to customise the
  report further in Word before distributing.
  Default: ~/Downloads/Telco_TI_Report.docx

TIPS
----
- You can leave entire sections blank — they will be skipped
- Save your YAML after filling — reload it for monthly updates
- The manifest YAML can also be loaded into Claude/GPT-4 to auto-fill
- Generate DOCX first, edit in Word, then share the PDF version

============================================================
