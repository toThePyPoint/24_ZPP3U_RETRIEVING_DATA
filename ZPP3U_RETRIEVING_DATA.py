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

from sap_conn import get_conn
from sap_rtab import rfc_read_table


if __name__ == "__main__":

    variant_name = "REP_BM_KPI_ALL"
    # variant_name = sys.argv[1]

    SAP_SYSTEM = "P11_SSO"
    # SAP_SYSTEM = "K11"

    BASE_PATH = Path(r"P:\Technisch\PLANY PRODUKCJI\PLANIŚCI\PP_TOOLS_TEMP_FILES\03_ZPP3U_RETRIEVING_DATA")
    ERROR_LOG_PATH = BASE_PATH / "error.log"

    delayed_btn_id = None

    start_time = time.time()

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

        if num_of_delayed_positions >= 5:
            zpp3u_scrolling = True
        else:
            zpp3u_scrolling = False
        result_dict = zpp3u_va03_get_data(sess, zpp3u_scrolling, num_of_delayed_positions)
        df = pd.DataFrame(result_dict)
        df['Description(PL)'] = ""
        df['Description(EN)'] = ""
        df['products_group'] = ""
        df['quantity_of_positions'] = 1

        df['customer_order'] = df['customer_order'].str.strip().str.zfill(10)

        # TODO: Implement the logic for retriving author and transport number with PyRFC
        customer_orders_list = df['customer_order'].to_list()

        if customer_orders_list:
            customer_ord_filter_vbeln = " OR ".join(
                [f"VBELN = '{num}'" for num in customer_orders_list]
            )

            customer_ord_filter_vbelv = " OR ".join(
                [f"VBELV = '{num}'" for num in customer_orders_list]
            )

            print(customer_ord_filter_vbeln)

            with get_conn(SAP_SYSTEM) as conn:

                vbak = rfc_read_table(
                    conn=conn,
                    table="VBAK",
                    fields=[
                        "VBELN", # customer order number
                        "ERNAM", # customer order author
                    ],
                    where=f"""
                        {customer_ord_filter_vbeln}
                    """,
                    # rowcount=1500
                )

                vbak_df = pd.DataFrame(vbak, columns=["VBELN", "ERNAM"])

                vbfa = rfc_read_table(
                    conn=conn,
                    table="VBFA",
                    fields=[
                        "VBELV",  # previous doc (SO)
                        "POSNV",  # pos SO
                        "VBELN",  # next doc
                        "POSNN",  # next doc pos
                        "VBTYP_N",  # doc type
                    ],
                    where=f"""
                        {customer_ord_filter_vbelv}
                    """
                )

                vbfa_df = pd.DataFrame(vbfa, columns=["VBELV", "POSNV", "VBELN", "POSNN", "VBTYP_N"])
        else:
            vbak_df = pd.DataFrame(columns=["VBELN", "ERNAM"])
            vbfa_df = pd.DataFrame(columns=["VBELV", "POSNV", "VBELN", "POSNN", "VBTYP_N"])

        # delivery_df = vbfa_df
        delivery_df = vbfa_df.loc[
            vbfa_df["VBTYP_N"] == "J"
            ]

        vbak_df.to_excel(fr"{BASE_PATH}\vbak_df.xlsx", index=False)
        delivery_df.to_excel(fr"{BASE_PATH}\delivery_df.xlsx", index=False)

        df = df.merge(delivery_df, left_on=['customer_order', 'customer_order_position'], right_on=['VBELV', 'POSNV'] , how='left')
        df.rename(columns={'VBELN': 'delivery_number', 'POSNN': 'delivery_position'}, inplace=True)

        df = df.merge(vbak_df, left_on=['customer_order'], right_on=['VBELN'], how='left')
        df.rename(columns={'ERNAM': 'creator'}, inplace=True)
        df.drop(columns=["VBELV", "VBELN", "POSNV"], inplace=True)

        print(vbak_df)
        print(delivery_df)
        print(df)
        df.to_excel(fr"{BASE_PATH}\main_df.xlsx", index=False)

        deliveries_list = delivery_df['VBELN'].str.strip().str.zfill(10).to_list()

        if deliveries_list:
            delivery_num_filter_vbeln = " OR ".join(
                [f"VBELN = '{num}'" for num in deliveries_list]
            )

            with get_conn(SAP_SYSTEM) as conn:

                vttp = rfc_read_table(
                    conn=conn,
                    table="VTTP",
                    fields=[
                        "VBELN",
                        "TKNUM",
                    ],
                    where=f"""
                        {delivery_num_filter_vbeln}
                    """
                )

            vttp_df = pd.DataFrame(vttp, columns=["VBELN", "TKNUM"])
        else:
            vttp_df = pd.DataFrame(columns=["VBELN", "TKNUM"])
        vttp_df.to_excel(fr"{BASE_PATH}\vttp_df.xlsx", index=False)

        df = df.merge(vttp_df, left_on=['delivery_number'], right_on=['VBELN'], how='left')
        df.rename(columns={'TKNUM': 'transport_number'}, inplace=True)
        df['transport_number'] = df['transport_number'].fillna("transport number not found")

        df["doc_date"] = pd.to_datetime(df["doc_date"], format="%d.%m.%Y").dt.strftime("%Y-%m-%d")
        df = df[["doc_date", "customer_order", "customer_order_position", "quantity_of_positions", "Description(PL)", "Description(EN)", "products_group", "creator", "transport_number"]]

        df_gr = df.groupby(['customer_order', 'transport_number'])['quantity_of_positions'].sum().reset_index()

        df = df.drop_duplicates(subset=['customer_order', 'transport_number'], keep='last')
        df = df.merge(df_gr, on=['customer_order', 'transport_number'], how='left')
        df = df.drop(columns=['quantity_of_positions_x'])
        df.rename(columns={'quantity_of_positions_y': 'quantity_of_positions_sum'}, inplace=True)
        df = df[['doc_date', 'customer_order', 'quantity_of_positions_sum', 'Description(PL)', 'Description(EN)',
                 'products_group', 'creator', 'transport_number']]

        # Convert DataFrame to clipboard-friendly format
        clipboard_data = df.to_csv(sep='\t', index=False, header=False)

        # Copy data to clipboard
        pyperclip.copy(clipboard_data)
        show_message("Dane skopiowane do schowka!")
        print(f"Total time of execution: {time.time() - start_time:.2f}")

    except Exception as e:
        print(e)
        logging.error("Error occurred", exc_info=True)
