import os
from pathlib import Path

# Paths
BASE_DIR = Path(__file__).parent.resolve()

# Browser & Profiles
BROWSER_CHANNEL = "chrome"  # Menggunakan Google Chrome yang sudah terinstal di Mac
BROWSER_DATA_DIR = BASE_DIR / "chrome_data"

# SSO BPS Credentials (dapat disesuaikan oleh pengguna lain)
SSO_USERNAME = os.getenv("SSO_USERNAME", "yudistiraelton")
SSO_PASSWORD = os.getenv("SSO_PASSWORD", "QP1al2zm3wo4sk5xn6")
AUTO_LOGIN_SSO = os.getenv("AUTO_LOGIN_SSO", "true").lower() in ("true", "1", "yes")

# URLs
FASIH_BASE_URL = "https://fasih-sm.bps.go.id"
SSO_LOGIN_URL = "https://fasih-sm.bps.go.id/oauth2/authorization/ics"
DEFAULT_ASSIGNMENT_URL = (
    "https://fasih-sm.bps.go.id/app/assignment/fd68e454-ba45-4b85-8205-f3bf777ded24/902b0ee4-4d3e-48ba-a40a-f4da93c46dae"
)

# Timeouts (in milliseconds)
DEFAULT_TIMEOUT = 30000
NAVIGATION_TIMEOUT = 45000

# Viewport
VIEWPORT = {"width": 1366, "height": 768}

