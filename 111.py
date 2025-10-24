import sys
import time
import winsound
import os
import json
import hashlib
import hmac
import subprocess
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QTextEdit, QLabel, QMessageBox, QFileDialog,
                             QProgressBar, QGroupBox, QFormLayout, QTimeEdit, QCheckBox,
                             QLineEdit, QGridLayout, QDialog, QDialogButtonBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QTime, QSettings
from PyQt5.QtGui import QColor
from PyQt5.QtGui import QIcon, QPixmap
import base64
from io import BytesIO
import requests
import urllib3
from openpyxl import load_workbook
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class LicenseManager:
    LICENSE_CODE = "*****"
    SECRET_KEY = "*****"
    SETTINGS_ORG = "HKTraffic"
    SETTINGS_APP = "HKTrafficLicense"
    TRIAL_FLAG_KEY = "trial_initialized"
    TRIAL_DEVICE_KEY = "trial_device_id"
    BACKWARD_TOLERANCE = timedelta(minutes=5)

    def __init__(self):
        self.settings = QSettings(self.SETTINGS_ORG, self.SETTINGS_APP)
        self.data = None
        self.device_serial = self._get_bios_serial()
        self.device_id = self._hash_serial(self.device_serial)
        self._load_data()
        if not self.data:
            self._ensure_trial_initialized()

    def _get_bios_serial(self):
        commands = [
            ["wmic", "bios", "get", "serialnumber"],
            ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_BIOS).SerialNumber"],
        ]
        for command in commands:
            try:
                output = subprocess.check_output(
                    command, stderr=subprocess.STDOUT, text=True, errors="ignore"
                )
                lines = [line.strip() for line in output.splitlines() if line.strip()]
                if not lines:
                    continue
                if lines[0].lower().startswith("serialnumber"):
                    lines = lines[1:]
                if lines:
                    return lines[0]
            except Exception:
                continue
        return "UNKNOWN-BIOS"

    def _hash_serial(self, serial):
        serial = (serial or "UNKNOWN-BIOS").strip()
        return hashlib.sha256(serial.encode("utf-8")).hexdigest()

    def _load_data(self):
        raw = self.settings.value("license_data", "")
        if not raw:
            self.data = None
            return
        try:
            data = json.loads(raw)
        except (TypeError, ValueError):
            self.data = None
            return
        if data.get("device_id") != self.device_id:
            self.data = None
            return
        if not self._is_signature_valid(data):
            self.data = None
            return
        data.pop("signature", None)
        self.data = data

    def _is_signature_valid(self, data):
        signature = data.get("signature")
        expected = self._signature(data)
        return signature == expected

    def _signature(self, data):
        payload = "|".join(
            [
                data.get("device_id", ""),
                data.get("activated_at", ""),
                data.get("expires_at", ""),
                data.get("last_check", ""),
                data.get("license_type", ""),
            ]
        )
        secret = self.SECRET_KEY.encode("utf-8")
        return hmac.new(secret, payload.encode("utf-8"), hashlib.sha256).hexdigest()

    def _persist_data(self):
        if not self.data:
            self.settings.remove("license_data")
            return
        data_to_store = dict(self.data)
        data_to_store["signature"] = self._signature(data_to_store)
        self.settings.setValue("license_data", json.dumps(data_to_store))
        self.settings.sync()

    def _trial_already_initialized(self):
        stored_flag = self.settings.value(self.TRIAL_FLAG_KEY, None)
        if isinstance(stored_flag, bool):
            flag = stored_flag
        elif stored_flag is None:
            flag = False
        else:
            flag = str(stored_flag).strip().lower() in ("true", "1", "yes")
        if not flag:
            return False
        stored_device = self.settings.value(self.TRIAL_DEVICE_KEY, "")
        if stored_device and stored_device != self.device_id:
            return False
        return True

    def _ensure_trial_initialized(self):
        if self._trial_already_initialized():
            return
        now = datetime.utcnow()
        expires = now + timedelta(days=30)
        self.data = {
            "device_id": self.device_id,
            "activated_at": now.isoformat(),
            "expires_at": expires.isoformat(),
            "last_check": now.isoformat(),
            "license_type": "trial",
        }
        self.settings.setValue(self.TRIAL_FLAG_KEY, True)
        self.settings.setValue(self.TRIAL_DEVICE_KEY, self.device_id)
        self._persist_data()
        self.settings.sync()

    def _parse_iso(self, value):
        if not value:
            return None
        try:
            return datetime.fromisoformat(value)
        except ValueError:
            return None

    def is_active(self):
        if not self.data:
            return False
        expires = self._parse_iso(self.data.get("expires_at"))
        if not expires:
            return False
        now = datetime.utcnow()
        return now <= expires

    def is_expired(self):
        if not self.data:
            return False
        return not self.is_active()

    def days_remaining(self):
        if not self.data:
            return 0
        expires = self._parse_iso(self.data.get("expires_at"))
        if not expires:
            return 0
        now = datetime.utcnow()
        if now >= expires:
            return 0
        remaining = expires - now
        days = remaining.days
        if remaining.seconds > 0:
            days += 1
        return max(days, 0)

    def ensure_valid(self):
        if not self.data:
            return False
        now = datetime.utcnow()
        expires = self._parse_iso(self.data.get("expires_at"))
        if not expires:
            self.clear()
            return False
        if now > expires:
            return False
        last_check = self._parse_iso(self.data.get("last_check"))
        if last_check and now < last_check - self.BACKWARD_TOLERANCE:
            self.clear()
            return False
        self.data["last_check"] = now.isoformat()
        self._persist_data()
        return True

    def activate(self, code):
        if not code or code.strip() != self.LICENSE_CODE:
            return False
        now = datetime.utcnow()
        expires = now + timedelta(days=30)
        self.data = {
            "device_id": self.device_id,
            "activated_at": now.isoformat(),
            "expires_at": expires.isoformat(),
            "last_check": now.isoformat(),
            "license_type": "activated",
        }
        self.settings.setValue(self.TRIAL_FLAG_KEY, True)
        self.settings.setValue(self.TRIAL_DEVICE_KEY, self.device_id)
        self._persist_data()
        self.settings.sync()
        return True

    def status_text(self):
        if not self.data:
            return "Not activated"
        license_type = self.data.get("license_type", "activated")
        if self.is_active():
            days = self.days_remaining()
            if license_type == "trial":
                return f"Trial active ({days} day(s) remaining)"
            return f"Activated ({days} day(s) remaining)"
        if license_type == "trial":
            return "Trial expired"
        return "License expired"

    def activated_on(self):
        if not self.data:
            return None
        return self.data.get("activated_at")

    def expires_on(self):
        if not self.data:
            return None
        return self.data.get("expires_at")

    def formatted_datetime(self, value):
        dt = self._parse_iso(value)
        if not dt:
            return None
        return dt.strftime("%Y-%m-%d %H:%M UTC")

    def formatted_activation(self):
        return self.formatted_datetime(self.activated_on())

    def formatted_expiry(self):
        return self.formatted_datetime(self.expires_on())

    def device_id_short(self):
        return self.device_id.upper()[:16]

    def license_type(self):
        if not self.data:
            return "none"
        return self.data.get("license_type", "unknown")

    def is_trial(self):
        return self.license_type() == "trial"

    def clear(self):
        self.data = None
        self.settings.remove("license_data")
        self.settings.sync()


def get_embedded_icon():
    icon_base64 = b"""
    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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAB4AAABaAAAAXAAAAB0AAAAAAAAAAAAAAAAAAAAAAAAAIwAAAHAAAABqAAAAKgAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABJAAAAvAAKBvUACQX4AAAAdAAAAAIAAAAAAAAAAAAAAAcAAACGABoP+wAXDfgAAADOAAAAYAAAABEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAcAAAAyAAAAngAMB+wAQiT+ALJi/wA+Iv8AAACgAAAADgAAAAAAAAAAAAAAEgAAAKoATCr/AL5p/wBXMP8AFQz2AAAAxAAAAE8AAAAPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAbAAAAeQACAdcAKBb/AIJI/wDkfv8A+Yr/AEQm/wAAAKMAAAAPAAAAAAAAAAAAAAATAAAArABQLP8A/43/APOG/wCdV/8ARSb/AAkF8QAAAJkAAABCAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAGAAAATQACAb8ACAX2AGc5/wDIb/8A9Yf/AP+N/wD/jf8ASSj/AAAApwAAABEAAAAAAAAAAAAAABMAAACsAFAs/wD/jf8A/43/APeJ/wDaeP8AkFD/ACIT+wAEAuMAAACCAAAAIwAAAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAGoAAADAABMK/QBfNf8Aw2z/APaI/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AO6E/wC3Zf8AVjD/AAQC8AAAAKsAAAA6AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAB0AAAB6AAMB2QAlFPwAgEf/ANZ2/wD7i/8A/o3/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP2M/wDlf/8AdED/AB0Q/AADAtEAAABdAAAAFAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAiAAAAgQAAAOAAHhD7AItN/wDTdf8A9If/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/APqK/wDogP8Aj0//ACoX/wABANoAAACFAAAAFQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFAAAAQAAAAJcACQX3ADwh/wCWU/8A8YX/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFAAAAIUAAADZAB8S/wCFSv8A4X3/APqK/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A+4v/ANBz/wB0Qf8AFAv4AAAA0AAAAFEAAAAJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABgAAADsAAACmAAcE7wBKKf8Aq1//AOqB/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AO6D/wClW/8AQST/AAQC8AAAAIUAAAAvAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAA0AAABKAAIBuQAQCfQAWDD/AMlv/wD0h/8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/APGF/wDAav8AYjb/AA4I8wABAbcAAABAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAcgAAAMkAGw//AG08/wDVdv8A+In/AP+N/wD/jf8A/43/AP+N/wD/jf8A/43/APqK/wDfe/8AekP/ACES/wAAANMAAABpAAAAFQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFAAAAKQAAAJIABQLoADce/gCfWP8A437/AP+N/wD/jf8A/43/AP+N/wDpgf8AsGH/AEMl/wAHBPAAAACWAAAANAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC0AAACiAAAA7QA/I/8Ar2H/AOuC/wDxhf8Aw2v/AF40/wAAAPMAAACzAAAAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4AAABfAAAAvgAXDf8AXzX/AHpE/wApF/8AAwLYAAAAcAAAABcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAAXAAAAeQAFA88ACgXgAAAAngAAACoAAAAFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEQAAACYAAAAGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA///////////////////////////////n5///////////////////g8H//////////////////4GA//////////////////wBgD/////////////////wAYAP////////////////wAGAA////////////////4ABgAH///////////////8AAYAAf//////////////4AAGAAB//////////////+AABgAAH/////////////8AAAYAAAf////////////+AAAGAAAD/////////////AAABgAAAP////////////wAAAYAAAA////////////8AAAGAAAAH////////////AAABgAAAAP///////////wAAAYAAAAB///////////8AAAGAAAAAP///////////AAABgAAAAAf////////h/wAAAYAAAAAD////////gP8AAAGAAAAAAH///////wD/AAABgAAAAAA///////wA/wAAAYAAAAAAH//////wAP8AAAGAAAAAAAP/////4AD/AAABgAAAAAAB/////8AA/wAAAYAAAAAAAP////4AAP8AAAGAAAAAAAAf///8AAD/AAABgAAAAAAAD///8AAA/wAAAYAAAAAAAAP//8AAAP8AAAGAAAAAAAAB//+AAAD/AAABgAAAAAAAAf//AAAA/wAAAYAAAAAAAAD//wAAAP8AAAGAAAAAAAAA//8AAAD/AAABgAAAAAAAAP//AAAA/wAAAYAAAAAAAAD//wAAAP8AAAGAAAAAAAAA//8AAAD/AAABgAAAAAAAAP//AAAA/wAAAYAAAAAAAAD//wAAAP8AAAGAAAAAAAAA//8AAAD/AAABgAAAgAAAAP//AAAA/wAAAYAAAMAAAAD//wAAAP8AAAGAAADwAAAA//8AAAD/AAABgAAA/AAAAP//AAAA/wAAAYAAAP4AAAD//wAAAP8AAAGAAAD/wAAB//8AAAD/AAABgAAA/+AAAf//AAAA/wAAAYAAAP/wAAH//wAAAP8AAAGAAAD//gAB//8AAAD+AAABgAAAP/8AAf//AAAA+AAAAYAAAB//4AH//4AAAPAAAAGAAAAH//AB//+AAABAAAABgAAAAf/8Af//gAAAAAAAAYAAAAD//wH//4AAAAAAAAGAAAAAP/+B//+AAAAAAAABgAAAAA//4///gAAAAAAAAYAAAAAH/////4AAAAAAAAGAAAAAAf////+AAAAAAAABgAAAAAB/////gAAAAAAAAYAAAAAAH////4AAAAAAAAGAAAAAAA////+AAAAAAAABgAAAAAAD////gAAAAAAAAYAAAAAAAP///4AAAAAAAAPgAAAAAAB///+AAAAAAAAGMAAAAAAAD///gAAAAAAAPBwAAAAAAAf//4AAAAAAAHAPAAAAAAAD//+AAAAAAACAAYAAAAAAAf//gAAAAAAHgADgAAAAAAD//4AAAAAADAAAOAAAAAAA//+AAAAAADgAAA4AAAAAAP//gAAAAADwAAAHAAAAAAD//4AAAAABwAAAAcAAAAAA//+AAAAABwAAAAB4AAAAAP//gAAAABwAAAAAGAAAAAD//4AAAAB4AAAAAA8AAAAA//+AAAAA8AAAAAADwAAAAP//gAAAA4AAAAAAAMAAAAD//4AAAA8AAAAAAAA4AAAA//+AAAAeAAAAAAAAHgAAAP//gAAAeAAAAAAAAAYAAAD//4AAAeAAAAAAAAABwAAA//+AAAOAAAAAAAAAAHAAAP//gAAPAAAAAAAAAAA8AAD//4AAPAAAAAAAAAAADgAA//+AAHAAAAAAAAAAAAOAAP//gAHgAAAAAAAAAAAB4AD//4AHgAAAAAAAAAAAADAA//+AHgAAAAAAAAAAAAAeAP//gDwAAAAAAAAAAAAADwD//4DwAAAAAAAAAAAAAAGA///D4AAAAAAAgAAAAAAA8///x4AAAAAAAcAAAAAAAB////4AAAAAAAf4AAAAAAAP///8AAAAAAAf/AAAAAAAD///+AAAAAAAP/4AAAAAAA////AAAAAAAP/8AAAAAAAP///wAAAAAAB/8AAAAAAAH////AAAAAAAP+AAAAAAAH////4AAAAAAB+AAAAAAAD/////gAAAAAACAAAAAAAB/////+AAAAAAAAAAAAAAD//////wAAAAAAAAAAAAAB///////gAAAAAAAAAAAAB///////8AAAAAAAAAAAAB////////gAAAAAAAAAAAA/////////AAAAAAAAAAAA/////////4AAAAAAAAAAA//////////wAAAAAAAAAA//////////+AAAAAAAAAAf//////////4AAAAAAAAAP///////////gAAAAAAAAf///////////8AAAAAAAAH////////////4AAAAAAAP/////////////AAAAAAAH/////////////4AAAAAAD//////////////wAAAAAH//////////////+AAAAAD///////////////8AAAAD////////////////gAAAD////////////////8AAAB/////////////////4AAB//////////////////AAA//////////////////+AB///////////////////wAf//////////////////+Af///////////////////8f//////////////////////////////8oAAAAQAAAAIAAAAABACAAAAAAAABAAAATCwAAEwsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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AAAAAAAAAAAAAAAAAAAAAAAOAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP////mf////////8A/////////AA////////wAA///////8AAA///////gAAA//////8AAAB//////wAAAB//////AAAAB/////8AAAAB////DwAAAAB///4PAAAAAD//+A8AAAAAD//gDwAAAAAD/8APAAAAAAH/AA8AAAAAAP8ADwAAAAAA/wAPAAAAAAD/AA8AAAAAAP8ADwAAAAAA/wAPAAAAAAD/AA8AAADAAP8ADwAAAOAA/wAPAAAA+AD/AA4AAAB+AP8ADAAAAB+A/wAAAAAAD+D/AAAAAAAD8P8AAAAAAAD//wAAAAAAAD//AAAAAAAAH/8AAAAAAAAH/wAAAAAAAAH/AAAAAAAAAP8AAAAAAAAA/wAABAAAAAD/AAAAAAAAAP8AAAAAAAAA/wAAAAAAgAD/AAIAAAAAAP8AAAAAAAAA/wAQAAAABAD/AEAAAAAAAP8AAAAAAAAA/wIAAAAAACD/CAAAAAAAAP+AAAAIAAAD/8AAAD4AAAP/wAAAfAAAA//gAAAwAAAH//AAAAAAAA///AAAAAAAP///AAAAAAD////AAAAAA/////AAAAAP////+AAAAB/////+AAAAP/////+AAAD//////+AAA///////+AAP///////8AB////////8Af////////8D/////////8/////KAAAADAAAABgAAAAAQAgAAAAAAAAJAAAEwsAABMLAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAMAAAAAAAAAAEAAAARAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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
     """
    pixmap = QPixmap()
    pixmap.loadFromData(base64.b64decode(icon_base64))
    return QIcon(pixmap)



class ScraperThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)
    progress_signal = pyqtSignal(int)

    def __init__(self, wifi_username, wifi_password, excel_path, parent=None):
        super().__init__(parent)
        self.wifi_username = wifi_username
        self.wifi_password = wifi_password
        self.excel_path = excel_path
        self.running = True


    def run(self):
        try:
            elements_to_cells = {
                "mapelement_23282_4": "B3",
                "mapelement_23282_10": "B4",
                "mapelement_23282_11": "B5",
                "mapelement_23282_12": "B6",
                "mapelement_23282_13": "B7",
                "mapelement_23282_23": "B11",
                "mapelement_23282_25": "B12",
                "mapelement_23282_24": "B13",
                "mapelement_23282_26": "B14",
                "mapelement_23282_14": "B17",
                "mapelement_23282_15": "B18",
                "mapelement_23282_16": "B20",
                "mapelement_23282_17": "B21",
                "mapelement_23282_18": "B24",
                "mapelement_23282_19": "B25",
                "mapelement_23282_27": "B26",
                "mapelement_23282_20": "B27",
                "mapelement_23282_53": "B2",
                "mapelement_23282_54": "B9",
                "mapelement_23282_55": "B10",
                "mapelement_23282_56": "B23",
                "mapelement_23282_21": "B28"
            }

            total_elements = len(elements_to_cells)
            current_progress = 0
            self.progress_signal.emit(5)

            self.update_signal.emit("Checking Excel file...")
            try:
                workbook = load_workbook(self.excel_path)
                sheet = workbook.active
            except Exception as e:
                self.finished_signal.emit(False, f"Error opening Excel file: {str(e)}")
                return

            self.progress_signal.emit(10)
            self.update_signal.emit("Initializing browser...")
            options = Options()
            options.add_argument("--headless")
            options.add_argument("--ignore-certificate-errors")
            driver = webdriver.Chrome(options=options)

            try:
                # Wifi Data Collection
                login_url = f"http://*****/login.htm?username={self.parent().wifi_user_edit.text().strip() if self.parent() else self.wifi_username}&password={self.parent().wifi_pass_edit.text() if self.parent() else self.wifi_password}"
                self.update_signal.emit(f"Logging in to Wifi via URL: http://*****/login.htm?username={self.wifi_username}&password=******")
                driver.get(login_url)
                time.sleep(5)
                self.progress_signal.emit(20)

                map_url = "http://*****/map.htm?id=23282&tabid=1"
                self.update_signal.emit("Navigating to wifi map page...")
                driver.get(map_url)
                time.sleep(5)
                self.progress_signal.emit(30)

                self.update_signal.emit("Reading wifi page data...")
                page_source = driver.page_source
                soup = BeautifulSoup(page_source, "html.parser")
                self.progress_signal.emit(40)

                alignment_center = Alignment(horizontal="center", vertical="center")
                font_size_16 = Font(size=16)

                self.update_signal.emit("Updating Excel file with wifi data...")
                # process first page (already in variable soup)
                for i, (element_id, cell) in enumerate(elements_to_cells.items()):
                    if not self.running:
                        break
                    div = soup.find("div", {"id": element_id})
                    if div:
                        label_spans = div.select("span.label")
                        if label_spans and len(label_spans) >= 3:
                            value = label_spans[2].text.strip()
                            self.update_signal.emit(f"{element_id} -> {value}")
                            sheet[cell] = value
                            sheet[cell].alignment = alignment_center
                            sheet[cell].font = font_size_16
                        else:
                            self.update_signal.emit(f"{element_id} -> Value not found")
                    else:
                        self.update_signal.emit(f"{element_id} -> Element not found")

                # second target: ***** id 3997
                if self.running:
                    try:
                        basra_ip    = self.parent().basra_ip_edit.text().strip()
                        basra_mapid = self.parent().basra_mapid_edit.text().strip()
                        basra_user  = self.parent().basra_user_edit.text().strip()
                        basra_pass  = self.parent().basra_pass_edit.text()
                        self.update_signal.emit(f"Logging into {basra_ip} ...")
                        driver.get(f"http://{basra_ip}/login.htm?username={basra_user}&password={basra_pass}")
                        time.sleep(2)
                        self.update_signal.emit(f"Navigating to {basra_ip} map {basra_mapid}...")
                        driver.get(f"http://{basra_ip}/map.htm?id={basra_mapid}&tabid=1")
                        time.sleep(3)
                        soup2 = BeautifulSoup(driver.page_source, "html.parser")
                        elements_3997 = {
                            f"mapelement_{basra_mapid}_1": "O14",
                            f"mapelement_{basra_mapid}_2": "O15",
                            f"mapelement_{basra_mapid}_4": "O16",
                            f"mapelement_{basra_mapid}_5": "O17",
                        }
                        for element_id, cell in elements_3997.items():
                            if not self.running:
                                break
                            div = soup2.find("div", {"id": element_id})
                            if div:
                                label_spans = div.select("span.label")
                                if label_spans and len(label_spans) >= 3:
                                    value = label_spans[2].text.strip()
                                    self.update_signal.emit(f"{element_id} -> {value}")
                                    sheet[cell] = value
                                    sheet[cell].alignment = alignment_center
                                    sheet[cell].font = font_size_16
                                else:
                                    self.update_signal.emit(f"{element_id} -> Value not found")
                            else:
                                self.update_signal.emit(f"{element_id} -> Element not found")
                    except Exception as e:
                        self.update_signal.emit(f"Error scraping *****: {e}")

                # third target: ***** id 3019
                if self.running:
                    try:
                        kut_ip    = self.parent().kut_ip_edit.text().strip()
                        kut_mapid = self.parent().kut_mapid_edit.text().strip()
                        kut_user  = self.parent().kut_user_edit.text().strip()
                        kut_pass  = self.parent().kut_pass_edit.text()
                        self.update_signal.emit(f"Logging into {kut_ip} ...")
                        driver.get(f"http://{kut_ip}/login.htm?username={kut_user}&password={kut_pass}")
                        time.sleep(2)
                        self.update_signal.emit(f"Navigating to {kut_ip} map {kut_mapid}...")
                        driver.get(f"http://{kut_ip}/map.htm?id={kut_mapid}&tabid=1")
                        time.sleep(3)
                        soup3 = BeautifulSoup(driver.page_source, "html.parser")
                        elements_3019 = {
                            f"mapelement_{kut_mapid}_15": "O9",
                            f"mapelement_{kut_mapid}_16": "O11",
                            f"mapelement_{kut_mapid}_18": "O10",
                            f"mapelement_{kut_mapid}_17": "O12",
                        }
                        for element_id, cell in elements_3019.items():
                            if not self.running:
                                break
                            div = soup3.find("div", {"id": element_id})
                            if div:
                                label_spans = div.select("span.label")
                                if label_spans and len(label_spans) >= 3:
                                    value = label_spans[2].text.strip()
                                    self.update_signal.emit(f"{element_id} -> {value}")
                                    sheet[cell] = value
                                    sheet[cell].alignment = alignment_center
                                    sheet[cell].font = font_size_16
                                else:
                                    self.update_signal.emit(f"{element_id} -> Value not found")
                            else:
                                self.update_signal.emit(f"{element_id} -> Element not found")
                    except Exception as e:
                        self.update_signal.emit(f"Error scraping *****: {e}")

                # Save
                if self.running:
                    try:
                        workbook.save(self.excel_path)
                        self.progress_signal.emit(100)
                        self.finished_signal.emit(True, "✅ Update completed successfully")
                    except PermissionError:
                        self.finished_signal.emit(False, "Error:...Could not save Excel file. Please close the file if it's open.")
                else:
                    self.finished_signal.emit(False, "Process stopped")
            finally:
                driver.quit()

        except Exception as e:
            self.finished_signal.emit(False, f"Unexpected error: {str(e)}")

    def stop(self):
        self.running = False
        self.update_signal.emit("Stopping process...")

class MainWindow(QMainWindow):


    def __init__(self):
        super().__init__()
        self.setWindowTitle("HK Traffic - NOC")
        self.setGeometry(100, 100, 1000, 700)
        
        # Dark theme colors
        self.dark_bg = QColor(30, 30, 30)
        self.green_text = QColor(0, 255, 0)
        self.dark_button = QColor(50, 50, 50)
        self.license_manager = LicenseManager()
        self._last_license_status = None
        
        self.excel_path = os.path.join(os.path.expanduser("~"), "Documents", "day.xlsx")
        self.wifi_username = "*****"
        self.wifi_password = "*****"
        self.scraper_thread = None
        
        self.init_ui()
        self.setup_scheduler()
        self.apply_dark_theme()
        self.refresh_license_state()
        
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout()
        central_widget.setLayout(main_layout)
        
        left_container = QWidget()
        left_layout = QVBoxLayout()
        left_container.setLayout(left_layout)
        
        right_container = QWidget()
        right_layout = QVBoxLayout()
        right_container.setLayout(right_layout)
        

        # ===== Basra FTTH PRTG =====
        self.basra_group = QGroupBox("Basra FTTH PRTG")
        basra_form = QFormLayout()
        self.basra_ip_edit = QLineEdit("*****")
        self.basra_mapid_edit = QLineEdit("3997")
        self.basra_user_edit = QLineEdit("*****")
        self.basra_pass_edit = QLineEdit("*****")
        self.basra_pass_edit.setEchoMode(QLineEdit.Password)
        basra_form.addRow(QLabel("IP:"), self.basra_ip_edit)
        basra_form.addRow(QLabel("Map ID:"), self.basra_mapid_edit)
        basra_form.addRow(QLabel("Username:"), self.basra_user_edit)
        basra_form.addRow(QLabel("Password:"), self.basra_pass_edit)
        self.basra_group.setLayout(basra_form)
        # Basra FTTH PRTG group hidden (not added to GUI)
# ===== Kut FTTH PRTG =====
        self.kut_group = QGroupBox("Kut FTTH PRTG")
        kut_form = QFormLayout()
        self.kut_ip_edit = QLineEdit("*****")
        self.kut_mapid_edit = QLineEdit("3019")
        self.kut_user_edit = QLineEdit("*****")
        self.kut_pass_edit = QLineEdit("*****")
        self.kut_pass_edit.setEchoMode(QLineEdit.Password)
        kut_form.addRow(QLabel("IP:"), self.kut_ip_edit)
        kut_form.addRow(QLabel("Map ID:"), self.kut_mapid_edit)
        kut_form.addRow(QLabel("Username:"), self.kut_user_edit)
        kut_form.addRow(QLabel("Password:"), self.kut_pass_edit)
        self.kut_group.setLayout(kut_form)
        # Kut FTTH PRTG group hidden (not added to GUI)
        main_layout.addWidget(left_container, stretch=3)
        main_layout.addWidget(right_container, stretch=1)

        title_label = QLabel("HK Traffic - Automated Data Collector - Wireless - FTTH NOC")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 20px;")
        left_layout.addWidget(title_label)
        
        self.file_label = QLabel(f"Excel file: {self.excel_path}")
        left_layout.addWidget(self.file_label)
        
        change_file_btn = QPushButton("Change Excel File")
        change_file_btn.clicked.connect(self.change_excel_file)
        left_layout.addWidget(change_file_btn)
        
        self.creds_group = QGroupBox("Wifi PRTG", self)
        self.creds_group.setParent(self)
        self.creds_group.hide()
        
        
        creds_layout = QFormLayout()
        
        self.wifi_user_edit = QLineEdit(self.wifi_username)
        self.wifi_pass_edit = QLineEdit(self.wifi_password)
        self.wifi_pass_edit.setEchoMode(QLineEdit.Password)
        
        creds_layout.addRow(QLabel("Wifi Username:"), self.wifi_user_edit)
        creds_layout.addRow(QLabel("Wifi Password:"), self.wifi_pass_edit)
        
        # Add editable IP & Map ID for the primary WiFi PRTG target
        self.wifi_ip_edit = QLineEdit("*****")
        self.wifi_mapid_edit = QLineEdit("23282")
        creds_layout.addRow(QLabel("IP:"), self.wifi_ip_edit)
        creds_layout.addRow(QLabel("Map ID:"), self.wifi_mapid_edit)
        save_creds_btn = QPushButton("Save Credentials")
        save_creds_btn.clicked.connect(self.save_credentials)
        left_layout.addWidget(save_creds_btn)
        
        self.creds_group.setLayout(creds_layout)
        # Wifi PRTG group hidden (not added to GUI)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        left_layout.addWidget(self.progress_bar)
        
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        left_layout.addWidget(self.log_area, stretch=1)
        
        buttons_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("Start")
        self.start_btn.clicked.connect(self.start_scraping)
        buttons_layout.addWidget(self.start_btn)
        
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_data)
        buttons_layout.addWidget(self.clear_btn)
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_scraping)
        buttons_layout.addWidget(self.stop_btn)
        
        about_btn = QPushButton("About")
        about_btn.clicked.connect(self.show_about)
        buttons_layout.addWidget(about_btn)
        
        left_layout.addLayout(buttons_layout)
        
        schedule_group = QGroupBox("Schedule Settings")
        schedule_group.setStyleSheet("QGroupBox { margin-top: 20px; }")
        schedule_layout = QVBoxLayout()
        
        self.time_checks = []
        self.time_edits = []
        
        default_times = [
            ("7:00 PM", QTime(18,55)),
            ("8:00 PM", QTime(20,55)),
            ("9:00 PM", QTime(21,55)),
            ("10:00 PM", QTime(22,55)),
            ("11:00 PM", QTime(23,55)),
            ("12:00 AM", QTime(0,55))
        ]
        
        for label, time in default_times:
            row_widget = QWidget()
            row_layout = QHBoxLayout()
            row_widget.setLayout(row_layout)
            
            check = QCheckBox()
            time_edit = QTimeEdit()
            time_edit.setDisplayFormat("hh:mm AP")
            time_edit.setTime(time)
            
            row_layout.addWidget(check)
            row_layout.addWidget(QLabel(label))
            row_layout.addWidget(time_edit)
            
            self.time_checks.append(check)
            self.time_edits.append(time_edit)
            
            schedule_layout.addWidget(row_widget)
        
        save_btn = QPushButton("Save Schedule")
        save_btn.clicked.connect(self.save_schedule)
        schedule_layout.addWidget(save_btn)
        
        schedule_group.setLayout(schedule_layout)
        right_layout.addWidget(schedule_group, stretch=1)
        
        right_layout.addStretch()

    def show_activation_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Activation")
        layout = QVBoxLayout(dialog)

        device_label = QLabel(f"Device ID: {self.license_manager.device_id_short()}")
        device_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(device_label)

        bios_value = self.license_manager.device_serial or "UNKNOWN-BIOS"
        bios_label = QLabel(f"BIOS Serial: {bios_value}")
        bios_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(bios_label)

        status_label = QLabel(f"Status: {self.license_manager.status_text()}")
        layout.addWidget(status_label)

        license_type = self.license_manager.license_type()
        friendly_type = {
            "trial": "Trial",
            "activated": "Activated",
            "none": "None",
        }.get(license_type, license_type.title())
        layout.addWidget(QLabel(f"License Type: {friendly_type}"))

        expiry_text = self.license_manager.formatted_expiry()
        if expiry_text:
            layout.addWidget(QLabel(f"Expires: {expiry_text}"))

        code_edit = QLineEdit()
        code_edit.setEchoMode(QLineEdit.Password)
        code_edit.setPlaceholderText("Enter activation code")
        layout.addWidget(code_edit)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(button_box)

        def handle_accept():
            code = code_edit.text().strip()
            if not code:
                QMessageBox.warning(dialog, "Activation", "Please enter the activation code.")
                return
            if self.license_manager.activate(code):
                QMessageBox.information(self, "Activation", "License activated successfully for 30 days.")
                self.refresh_license_state()
                dialog.accept()
            else:
                QMessageBox.critical(dialog, "Activation", "Invalid activation code for this device.")

        button_box.accepted.connect(handle_accept)
        button_box.rejected.connect(dialog.reject)
        code_edit.returnPressed.connect(handle_accept)

        dialog.exec_()

    def refresh_license_state(self):
        is_active = self.license_manager.is_active()
        is_running = self.scraper_thread and self.scraper_thread.isRunning()
        status_text = self.license_manager.status_text()
        device_id = self.license_manager.device_id_short()
        license_type = self.license_manager.license_type()
        friendly_type = {
            "trial": "Trial",
            "activated": "Activated",
            "none": "None",
        }.get(license_type, license_type.title())
        tooltip_lines = [
            f"Status: {status_text}",
            f"License: {friendly_type}",
            f"Device ID: {device_id}",
        ]
        expiry_text = self.license_manager.formatted_expiry()
        if is_active and expiry_text:
            tooltip_lines.append(f"Expires: {expiry_text}")
        else:
            if self.license_manager.is_trial():
                tooltip_lines.append("Trial expired. Enter activation code to continue.")
            else:
                tooltip_lines.append("Activate with your code to enable Start.")
        tooltip = "\n".join(tooltip_lines)

        if hasattr(self, "start_btn"):
            self.start_btn.setEnabled(is_active and not is_running)
            self.start_btn.setToolTip("" if is_active else tooltip)

        if self._last_license_status != status_text and hasattr(self, "log_area"):
            self.log(f"License status: {status_text}")
            self._last_license_status = status_text

    def save_credentials(self):
        self.wifi_username = self.wifi_user_edit.text()
        self.wifi_password = self.wifi_pass_edit.text()
        
        self.log("Credentials saved successfully")
        QMessageBox.information(self, "Success", "Credentials saved successfully!")
    
    def setup_scheduler(self):
        self.schedule_timer = QTimer(self)
        self.schedule_timer.timeout.connect(self.check_schedule)
        self.schedule_timer.start(60000)
        self.schedule_times = []
        self.load_schedule()
    
    def load_schedule(self):
        default_times = [
            QTime(19, 0), QTime(20, 0), QTime(21, 0),
            QTime(22, 0), QTime(23, 0), QTime(0, 0)
        ]
        self.schedule_times = default_times
        self.update_schedule_ui()
    
    def save_schedule(self):
        self.schedule_times = []
        for i in range(len(self.time_checks)):
            if self.time_checks[i].isChecked():
                self.schedule_times.append(self.time_edits[i].time())
        self.log("Schedule saved successfully")
        QMessageBox.information(self, "Success", "Schedule settings saved successfully!")
    
    def update_schedule_ui(self):
        for i in range(len(self.time_checks)):
            time = self.time_edits[i].time()
            self.time_checks[i].setChecked(time in self.schedule_times)
    
    def check_schedule(self):
        self.refresh_license_state()
        current_time = QTime.currentTime()
        for scheduled_time in self.schedule_times:
            if (current_time.hour() == scheduled_time.hour() and 
                current_time.minute() == scheduled_time.minute()):
                if not hasattr(self, 'last_triggered') or self.last_triggered != scheduled_time:
                    self.last_triggered = scheduled_time
                    self.trigger_scheduled_task(scheduled_time)
                    break
    
    def trigger_scheduled_task(self, time):
        # Validate license before performing any scheduled actions
        if not self.license_manager.ensure_valid():
            status = self.license_manager.status_text()
            # Only notify that the license has expired; do not start
            self.log(f"Scheduled run skipped. {status}.")
            try:
                QMessageBox.warning(
                    self,
                    "License Expired",
                    f"Scheduled run skipped. License is {status}."
                )
            except Exception:
                pass
            self.refresh_license_state()
            return
        self.log(f"⏰ Scheduled task triggered for {time.toString('hh:mm AP')}")
        self.play_alert()
        self.show_notification(time)
        self.start_scraping()





    
    def stop_scraping(self):
        if hasattr(self, "scraper_thread") and self.scraper_thread and self.scraper_thread.isRunning():
            self.scraper_thread.stop()
            self.log("Stopping process...")
        else:
            self.log("No running process to stop.")

    def play_alert(self):
        try:
            winsound.Beep(1000, 500)
        except:
            self.log("⚠️ Could not play alert sound")
    
    def show_notification(self, time):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Notification")
        msg.setText(f"Time to collect data!\nTime: {time.toString('hh:mm AP')}")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
    
    def apply_dark_theme(self):
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {self.dark_bg.name()};
                color: {self.green_text.name()};
            }}
            QPushButton {{
                background-color: {self.dark_button.name()};
                color: {self.green_text.name()};
                border: 1px solid {self.green_text.name()};
                padding: 5px;
            }}
            QPushButton:hover {{
                background-color: {self.green_text.name()};
                color: black;
            }}
            QTextEdit {{
                background-color: black;
                color: {self.green_text.name()};
                font-family: Consolas, monospace;
            }}
            QLabel {{
                color: {self.green_text.name()};
            }}
            QGroupBox {{
                border: 1px solid {self.green_text.name()};
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 15px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }}
            QProgressBar {{
                border: 1px solid {self.green_text.name()};
                text-align: center;
                height: 20px;
            }}
            QProgressBar::chunk {{
                background-color: {self.green_text.name()};
            }}
            QLineEdit {{
                background-color: black;
                color: {self.green_text.name()};
                border: 1px solid {self.green_text.name()};
            }}
        """)
        
    def change_excel_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excel_path = file_path
            self.file_label.setText(f"Excel file: {self.excel_path}")
            self.log(f"Excel file changed to: {file_path}")
    
    def start_scraping(self):
        if self.scraper_thread and self.scraper_thread.isRunning():
            QMessageBox.warning(self, "Warning", "Process already running!")
            return
        
        if not self.license_manager.ensure_valid():
            status = self.license_manager.status_text()
            self.log(f"Start aborted. License status: {status}.")
            QMessageBox.warning(
                self,
                "Activation Required",
                f"License is {status}. Please activate with your code to enable Start."
            )
            self.refresh_license_state()
            return
        
        try:
            with open(self.excel_path, 'a') as f:
                pass
        except IOError:
            QMessageBox.critical(self, "Error", "Excel file is locked. Please close it in other programs.")
            return
        
        try:
            load_workbook(self.excel_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Cannot open Excel file:\n{str(e)}")
            return
        


        self.log("Starting data collection...")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.clear_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        # Refresh Wifi PRTG creds from UI before starting thread
        self.wifi_username = self.wifi_user_edit.text().strip()
        self.wifi_password = self.wifi_pass_edit.text()
        self.scraper_thread = ScraperThread(self.wifi_username, self.wifi_password, self.excel_path, parent=self)
        self.scraper_thread.update_signal.connect(self.log)
        self.scraper_thread.finished_signal.connect(self.on_scraping_finished)
        self.scraper_thread.progress_signal.connect(self.update_progress)
        self.scraper_thread.start()
        

#.................................................................................................................
                



    def update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def on_scraping_finished(self, success, message):
        self.stop_btn.setEnabled(False)
        self.clear_btn.setEnabled(True)
        self.scraper_thread = None
        self.refresh_license_state()
        self.log(message)
        if success:
            QMessageBox.information(self, "Success", "Excel file updated successfully!")
        else:
            QMessageBox.warning(self, "Warning", message)
    
    def clear_data(self):
        reply = QMessageBox.question(self, "Confirmation", "Clear log and reset Excel file?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.log_area.clear()
            self.progress_bar.setValue(0)
            self.log("Log cleared and progress reset.")
            
            try:
                workbook = load_workbook(self.excel_path)
                sheet = workbook.active
                
                cells_to_clear = ["B3", "B2", "B9", "B10", "B23", "B4", "B5", "B6", "B7", "B11", "B12", "B13", "B14",
                                "B17", "B18", "B20", "B21", "B24", "B25", "B26", "B16", "B27", "B28", "O9", "O10", "O11", "O12", "O14", "O15", "O16", "O17"]
                
                for cell in cells_to_clear:
                    sheet[cell] = ""
                
                workbook.save(self.excel_path)
                self.log("Excel data cleared.")
            except Exception as e:
                self.log(f"Error clearing Excel: {str(e)}")
                QMessageBox.warning(self, "Error", f"Cannot clear Excel file:\n{str(e)}")
    
    def show_about(self):
        self.refresh_license_state()
        dialog = QDialog(self)
        dialog.setWindowTitle("About HK Traffic")
        layout = QVBoxLayout(dialog)

        info_label = QLabel()
        info_label.setTextFormat(Qt.RichText)
        info_label.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.LinksAccessibleByMouse)
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        activate_btn = QPushButton()
        button_box.addButton(activate_btn, QDialogButtonBox.ActionRole)
        layout.addWidget(button_box)

        def update_contents():
            info_label.setText(self._build_about_html())
            if self.license_manager.is_active():
                activate_btn.setText("Activate" if self.license_manager.is_trial() else "Re-Activate")
            else:
                activate_btn.setText("Activate")

        def handle_activate():
            self.show_activation_dialog()
            self.refresh_license_state()
            update_contents()

        activate_btn.clicked.connect(handle_activate)
        close_button = button_box.button(QDialogButtonBox.Close)
        if close_button:
            close_button.clicked.connect(dialog.accept)
        else:
            button_box.rejected.connect(dialog.reject)

        update_contents()
        dialog.exec_()

    def _build_about_html(self):
        license_status = self.license_manager.status_text()
        days_remaining = self.license_manager.days_remaining()
        activated_text = self.license_manager.formatted_activation() or "N/A"
        expiry_text = self.license_manager.formatted_expiry() or "N/A"
        device_id = self.license_manager.device_id_short()
        bios_serial = self.license_manager.device_serial or "UNKNOWN-BIOS"
        license_type = self.license_manager.license_type()
        friendly_type = {
            "trial": "Trial",
            "activated": "Activated",
            "none": "None",
        }.get(license_type, license_type.title())
        return f"""
        <h2>HK Traffic</h2>
        <p>Version 5.1</p>
        <p>HK Traffic - Wireless - FTTH NOC</p>
        <p>- Automated Data Collector -</p>
        <p>Copyright Ac 2025 *****.</p>
        <p># Hyper Knowledge</p>
        <p>Features:</p>
        <ul>
            <li>Automated Wifi Data Collection</li>
            <li>Credentials Management</li>
            <li>Scheduling System</li>
            <li>Progress Tracking</li>
            <li>Developed with PyQt5 and Selenium</li>
            <li>Improved GUI</li>
        </ul>
        <hr>
        <p><b>License Status:</b> {license_status}</p>
        <p><b>License Type:</b> {friendly_type}</p>
        <p><b>Days Remaining:</b> {days_remaining}</p>
        <p><b>Activated:</b> {activated_text}</p>
        <p><b>Expires:</b> {expiry_text}</p>
        <p><b>Device ID:</b> {device_id}</p>
        <p><b>BIOS Serial:</b> {bios_serial}</p>
        """
    def log(self, message):
        if isinstance(message, str) and "mapelement_" in message:
            return
        self.log_area.append(message)
        self.log_area.verticalScrollBar().setValue(self.log_area.verticalScrollBar().maximum())
    
    def closeEvent(self, event):
        if self.scraper_thread and self.scraper_thread.isRunning():
            reply = QMessageBox.question(self, "Process Running",
                                       "Process is still running, stop and exit?",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.scraper_thread.stop()
                self.scraper_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()



if __name__ == "__main__":
    app = QApplication(sys.argv)

    
    app.setWindowIcon(get_embedded_icon())

    window = MainWindow()
    window.setWindowIcon(get_embedded_icon())

    window.show()
    sys.exit(app.exec_())
