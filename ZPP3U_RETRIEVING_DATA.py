import os
import time
import sys
import ctypes
from datetime import datetime
from pathlib import Path
import logging

import pandas as pd

from sap_connection import get_last_session, get_client


if __name__ == "__main__":

    # Get today's date in YYYYMMDD format
    today = datetime.today().strftime("%Y_%m_%d")
    start_time = datetime.now().strftime("%H:%M:%S")
    username = os.getlogin()

    BASE_PATH = Path(r"P:\Technisch\PLANY PRODUKCJI\PLANIŚCI\PP_TOOLS_TEMP_FILES\03_ZPP3U_RETRIEVING_DATA")
    ERROR_LOG_PATH = BASE_PATH / "error.log"

    # Hide console window
    if sys.platform == "win32":
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        hWnd = kernel32.GetConsoleWindow()
        if hWnd:
            user32.ShowWindow(hWnd, 6)  # 6 = Minimize

    logging.basicConfig(
        filename=ERROR_LOG_PATH,
        level=logging.ERROR,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )

    try:
        sess, tr, nu = get_last_session(max_num_of_sessions=6)

    except Exception as e:
        logging.error("Error occurred", exc_info=True)
