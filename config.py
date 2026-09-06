import os

from dotenv import load_dotenv


# ============================================================
# LOAD ENVIRONMENT VARIABLES
# ============================================================

load_dotenv()


# ============================================================
# APPLICATION MODE
# ============================================================

# "work"   = normal work network
# "remote" = home / remote RPii connection

MODE = "remote"


# ============================================================
# FILES / FOLDERS
# ============================================================

ORDERS_FILE = "orders.xlsx"

INVOICE_FOLDER = (
    r"C:\Users\Noel\Documents\OK TO DELETE"
)


# ============================================================
# RPii
# ============================================================

WORK_URL = (
    "https://gar.rpdns.co.uk/"
    "companyGAR/rp2.cgi"
)

REMOTE_URL = (
    "https://garremote.rpdns.co.uk/"
    "companyGAR/rp2.cgi"
)

RP2_REMOTE_USERNAME = os.getenv(
    "RP2_REMOTE_USERNAME"
)

RP2_REMOTE_PASSWORD = os.getenv(
    "RP2_REMOTE_PASSWORD"
)


if MODE == "remote":

    RP2_URL = REMOTE_URL

    if (
        not RP2_REMOTE_USERNAME
        or not RP2_REMOTE_PASSWORD
    ):
        raise RuntimeError(
            "Remote RPii username/password "
            "missing from .env file."
        )

elif MODE == "work":

    RP2_URL = WORK_URL

else:

    raise ValueError(
        'MODE must be either "work" or "remote".'
    )


# ============================================================
# MIRAKL / DIY
# ============================================================

MIRAKL_BASE_URL = os.getenv(
    "MIRAKL_BASE_URL"
)

MIRAKL_API_KEY = os.getenv(
    "MIRAKL_API_KEY"
)


if not MIRAKL_BASE_URL:
    raise RuntimeError(
        "MIRAKL_BASE_URL missing from .env"
    )

if not MIRAKL_API_KEY:
    raise RuntimeError(
        "MIRAKL_API_KEY missing from .env"
    )


# ============================================================
# MIRAKL DIY SETTINGS
# ============================================================

DIY_DELIVERY_TOPIC_CODE = "44"

LIVE_MIRAKL_SEND_ENABLED = False    

DIY_DIRECT_DELIVERY_CARRIER_CODE = "DIR"

DIY_DIRECT_DELIVERY_CARRIER_LABEL = (
    "The verified seller will contact you directly "
    "to plan a delivery"
)