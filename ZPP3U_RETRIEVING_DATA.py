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
from sap_transactions import partial_matching, zpp3u_va03_get_data
from gui_manager import show_message


if __name__ == "__main__":

    # variant_name = "REP_LU_KPI_ALL"
    variant_name = sys.argv[1]

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
        simple_load_variant(sess, variant_name, True)
        # Pause to select date
        show_message("Wybierz daty w transakcji ZPP3U i kliknij 'OK'.")
        sess.findById("wnd[0]").sendVKey(8)

        element_id = None

        for i in range(1, 500, 1):

            if not element_id:
                element_id = partial_matching(sess, rf"lbl\[64,{str(i-1)}\]")

            if element_id:
                element_id = str.replace(element_id, f',{i-1}]', f',{i}]')
                # sess.findById(element_id).setFocus()
                # '/app/con[0]/ses[2]/wnd[0]/usr/lbl[94,18]'
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
            time.sleep(2)
        else:
            show_message("There is no data.")

        result_dict = zpp3u_va03_get_data(sess)
        df = pd.DataFrame(result_dict)
        df['Description(PL)'] = ""
        df['Description(EN)'] = ""
        df['products_group'] = ""
        df['quantity_of_positions'] = 1

        df["doc_date"] = pd.to_datetime(df["doc_date"], format="%d.%m.%Y").dt.strftime("%Y-%m-%d")
        df = df[["doc_date", "customer_order", "quantity_of_positions", "Description(PL)", "Description(EN)", "products_group", "creator"]]

        df_gr = df.groupby('customer_order')['quantity_of_positions'].sum()
        df_gr = df_gr.reset_index()

        df = df.drop_duplicates(subset=['customer_order'], keep='last')
        df = df.merge(df_gr, on='customer_order', how='left')
        df = df.drop(columns=['quantity_of_positions_x'])
        df.rename(columns={'quantity_of_positions_y': 'quantity_of_positions_sum'}, inplace=True)
        df = df[['doc_date', 'customer_order', 'quantity_of_positions_sum', 'Description(PL)', 'Description(EN)',
                 'products_group', 'creator']]

        # Convert DataFrame to clipboard-friendly format
        clipboard_data = df.to_csv(sep='\t', index=False, header=False)

        # Copy data to clipboard
        pyperclip.copy(clipboard_data)

    except Exception as e:
        print(e)
        logging.error("Error occurred", exc_info=True)
