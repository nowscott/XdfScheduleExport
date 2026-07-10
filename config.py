"""项目配置。"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).parent
USER_DATA_DIR = Path(os.getenv("XDF_USER_DATA_DIR", BASE_DIR / "browser_data"))
EXPORT_DIR = Path(os.getenv("XDF_EXPORT_DIR", BASE_DIR / "output"))

SCHEDULE_PAGE_URL = os.getenv(
    "XDF_SCHEDULE_PAGE_URL",
    "https://we.xdf.cn/easyServer/mySchedule/myScheduleList",
)
API_BASE = os.getenv(
    "XDF_API_BASE",
    "https://gw-xeasy.xdf.cn/xeasy-srv-teachinghub",
)

LOGIN_TIMEOUT = int(os.getenv("XDF_LOGIN_TIMEOUT", "180"))
REQUEST_INTERVAL = float(os.getenv("XDF_REQUEST_INTERVAL", "0.4"))

EXPORT_DIR.mkdir(parents=True, exist_ok=True)
