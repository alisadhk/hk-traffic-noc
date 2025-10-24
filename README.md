<h1 align="center">HK Traffic - NOC</h1>

<p align="center"><b>Collects traffic graphs from PRTG maps for NOC reporting</b></p>

<p align="center">
  ![Python](https://img.shields.io/badge/Python-3.12-3776AB?logo=python&logoColor=white)
  ![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)
  ![Updated](https://img.shields.io/badge/updated-2025-10-24-success)
  [![Releases](https://img.shields.io/badge/downloads-Releases-blue)](https://github.com/alisadhk/hk-traffic-noc/releases/latest)
</p>

---

## Overview
A PyQt5 desktop tool for Network Operations Centers. It logs in to PRTG web maps with Selenium for several sites (FTTH and Wi-Fi), reads the traffic values and writes them into an Excel report on a schedule, with sound and desktop notifications.

## Features
- Separate profiles for Basra FTTH, Kut FTTH and Wi-Fi PRTG servers
- Scheduled collection (Schedule Settings)
- Excel output file you can change
- Saved credentials (Save Credentials)
- Start / Stop / Clear controls with a live log
- Activation dialog (trial/licensed mode)

## Screenshots
<p align="center"><img src="screenshots/screenshot-1.png" alt="HK Traffic - NOC screenshot" width="720"></p>

## Download and install
### Option A - Ready-to-run download (recommended)
1. Open the **[Releases page](https://github.com/alisadhk/hk-traffic-noc/releases/latest)**.
2. Download the file(s) listed under **Assets**:

| File | Size |
|---|---|
| `HK-Traffic-NOC.exe` | 71.6 MB |

3. Run it. 

### Option B - Run from source
Requirements: Windows 10/11 and Python 3.10+.

```bash
git clone https://github.com/alisadhk/hk-traffic-noc.git
cd hk-traffic-noc
python -m pip install -r requirements.txt
python 111.py
```

## How to use
1. Start the program.
2. For each site enter the PRTG **IP**, **Map ID**, **Username** and **Password** (or the Wi-Fi username/password).
3. Pick the Excel file with **Change Excel File**.
4. Optionally set the schedule in **Schedule Settings** and press **Save Schedule**.
5. Press **Start**. Watch the log; press **Stop** to end the run.

## Configuration and privacy
- All sensitive values from the original private version (passwords, tokens, license codes, user names, internal IP addresses, personal names and links) were replaced with `*****`.
- The built-in license code and signing key are masked (`*****`). Put your own values in 111.py (LicenseManager) before building a licensed version, or use the trial mode.
- Default IP/user/password fields are empty placeholders; type your own values in the UI.

## Build the executable yourself
```bash
python -m pip install pyinstaller
python -m PyInstaller --noconfirm 111.spec
```
The executable is created in the `dist/` folder.

## Notes
- Needs Google Chrome installed (Selenium Manager downloads the matching driver).
- Network maps, IDs and addresses from the original deployment are not published.

---

<p align="center">Maintained by <a href="https://github.com/alisadhk">@alisadhk</a> - last update 2025-10-24</p>