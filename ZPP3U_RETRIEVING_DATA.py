import os
import time
import sys
import ctypes
from datetime import datetime
from pathlib import Path
import logging

import pandas as pd

from sap_connection import get_last_session
from sap_functions import open_one_transaction, simple_load_variant
from sap_transactions import partial_matching
from gui_manager import show_message


if __name__ == "__main__":

    variant_name = "REP_LU_KPI_ALL"

    BASE_PATH = Path(r"P:\Technisch\PLANY PRODUKCJI\PLANIŚCI\PP_TOOLS_TEMP_FILES\03_ZPP3U_RETRIEVING_DATA")
    ERROR_LOG_PATH = BASE_PATH / "error.log"

    delayed_btn_id = None

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
        open_one_transaction(sess, "ZPP3U")
        simple_load_variant(sess, variant_name)

        for i in range(1, 500, 1):
            # print(sess.findById(f"wnd[0]/usr/lbl[64,{str(i)}]").text)
            element_id = partial_matching(sess, rf"lbl\[64,{str(i)}\]")
            if element_id:
                # sess.findById(element_id).setFocus()
                helper_field_id = partial_matching(sess, rf"lbl\[94,{str(i)}\]")
                delayed_btn_id = element_id
                if helper_field_id:
                    if sess.findById(helper_field_id).text == "ZA PÓŹNO":
                        break

        # Get to delayed orders list
        num_of_delayed_positions = sess.findById(delayed_btn_id).text
        num_of_delayed_positions = int(num_of_delayed_positions.strip())
        if num_of_delayed_positions > 0:
            sess.findById(delayed_btn_id).setFocus()
            sess.findById("wnd[0]").sendVKey(2)
        else:
            show_message("There is no data.")

    except Exception as e:
        print(e)
        logging.error("Error occurred", exc_info=True)
