# configuration module

import os
from dotenv import load_dotenv

load_dotenv()

BASE_URL = "https://wceasy.club/staff/"
TABLE_VIEW_URL = "https://wceasy.club/staff/table-view.php"
VIEW_EXCEL_BASE_URL = "https://wceasy.club/staff/view-excel.php?id="

WAIT_TIME = int(os.getenv("WAIT_TIME", "5"))


