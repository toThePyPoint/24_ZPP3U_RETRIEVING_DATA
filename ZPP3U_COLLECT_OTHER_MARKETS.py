import os
import time
import sys
import ctypes
from datetime import datetime
from pathlib import Path
import logging
import pyperclip

import pandas as pd

from sap_connection import get_last_session
from sap_functions import open_one_transaction, simple_load_variant
from sap_transactions import partial_matching
from gui_manager import show_message, OptionSelector
from other_functions import get_last_n_working_days

if __name__ == "__main__":

    variants_list = [
        "REP_LU_KPI_ALL",
        "REP_LU_KPI_AT",
        "REP_LU_KPI_BNL",
        "REP_LU_KPI_CHZ",
        "REP_LU_KPI_CZ",
        "REP_LU_KPI_DE",
        "REP_LU_KPI_FR",
        "REP_LU_KPI_HU",
        "REP_LU_KPI_IT",
        "REP_LU_KPI_PL"
    ]

    result_dict = dict()

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

    # Implement date selection
    selection_dates = get_last_n_working_days(12)
    selector = OptionSelector(selection_dates, "Select a date.")
    users_date = selector.show()

    try:
        sess, tr, nu = get_last_session(max_num_of_sessions=6)

        for variant_name in variants_list:
            total_sum_btn_id = None
            open_one_transaction(sess, "ZPP3U")
            simple_load_variant(sess, variant_name, True)
            # Implement dates insertion
            sess.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = users_date
            # sess.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = users_date
            sess.findById("wnd[0]").sendVKey(8)

            sys_msg_id = partial_matching(sess, rf"lbl\[0,{str(6)}\]")
            if sys_msg_id:
                sys_msg = sess.findById(sys_msg_id).text
                if sys_msg == "Nie znaleziono danych":
                    result_dict.setdefault(variant_name, []).append("0")
                    continue

            element_id = None

            for i in range(1, 500, 1):

                if not element_id:
                    element_id = partial_matching(sess, rf"lbl\[4,{str(i - 1)}\]")

                if element_id:
                    element_id = str.replace(element_id, f',{i - 1}]', f',{i}]')
                    # sess.findById(element_id).setFocus()
                    # '/app/con[0]/ses[2]/wnd[0]/usr/lbl[94,18]'
                    helper_field_id = partial_matching(sess, rf"lbl\[94,{str(i)}\]")
                    total_sum_btn_id = element_id
                    if helper_field_id:
                        if sess.findById(helper_field_id).text == "ZA PÓŹNO":
                            break

            # Get to delayed orders list
            total_sum_of_positions = sess.findById(total_sum_btn_id).text
            total_sum_of_positions = int(total_sum_of_positions.strip())

            result_dict.setdefault(variant_name, []).append(total_sum_of_positions)

        df = pd.DataFrame(result_dict)

        # Convert DataFrame to clipboard-friendly format
        clipboard_data = df.to_csv(sep='\t', index=False, header=False)

        # Copy data to clipboard
        pyperclip.copy(clipboard_data)
        show_message("Dane skopiowane do schowka!")

    except Exception as e:
        print(e)
        logging.error("Error occurred", exc_info=True)
