import datetime
import logging
import os
import sys
import time
import pandas as pd
import xlwings as xw
from kiteconnect import KiteConnect

logging.basicConfig(filename='mylog', level=logging.DEBUG)
log = logging.getLogger(__name__)


def get_instrument_list(kite,segment=None, month_count=None, year_of_expiry=None, list_of_instruments=None):
    try:
        if month_count is None:
            month_count = 12
        if year_of_expiry is None:
            year_of_expiry = 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030

        while True:
            try:
                master_contract = pd.DataFrame(kite.instruments('NFO'))
                break
            except Exception as e:
                logging.info("REST API to get Instrument failed. Retry in 1 sec !!!".format(e))
                time.sleep(1)

        # Remove the columns which is not required from data dictionary. axis=1 meant for column deletion.
        master_contract = master_contract.drop(["instrument_token", "exchange_token", "last_price", "tick_size"],
                                               axis=1)

        # Drop other than NSE exchange instruments in the row and remove other than Nifty,banknifty and finnifty
        # master_contract = master_contract.query("segment == 'NFO-OPT'")
        # master_contract = master_contract.query("name == ['NIFTY','BANKNIFTY','FINNIFTY']")
        if segment is not None:
            master_contract = master_contract.query('segment == @segment')
        if list_of_instruments is not None:
            master_contract = master_contract.query("name == @list_of_instruments")

        # Filter only current month and next month expiry
        df1 = pd.DataFrame()
        df1["expiry"] = pd.to_datetime(master_contract['expiry'])
        master_contract = master_contract[(df1['expiry'].dt.year.isin(year_of_expiry))
                                          & ((df1['expiry'].dt.month >= datetime.date.today().month)
                                             & (df1['expiry'].dt.month <= datetime.date.today().month + month_count))]
        df1["expiry"] = None
        master_contract["watchlist_symbol"] = master_contract["exchange"] + ":" + master_contract["tradingsymbol"]

        master_contract.columns = master_contract.columns.str.replace("_", " ")
        master_contract.columns = master_contract.columns.str.title()
        logging.info(str(master_contract.shape[0]) + ' rows of instruments is pulled and saved in exchange sheet!!!')
        # break
    except Exception as e:
        logging.info('Oops!!! Error occured on getting instruments symbols in exchange sheet !!!')
        logging.info("Excel creation failed: {}".format(e))
        raise Exception(e)
        # time.sleep(1)
    return master_contract


if __name__ == '__main__':
    kite = KiteConnect(enc_token=
                       "Hp5gH03rnMV1y2zrqtcxmupnfElj0cxzgCMZmtMYvsiCQOilAQsSfqOGLYcPLXXucLKJ8OLo06AIfypvVxK6BnduRaIR44P8+xLoAITvQNT2Fp56WPNXYw==")
    excel_workbook_testing = 'F:\SIMBU\Python\Trading\Zerodha\TradeOn_FromGetInstrumentList.xlsx'

    if not os.path.exists(excel_workbook_testing):
        try:
            wb_test = xw.Book()
            wb_test.save(excel_workbook_testing)
            wb_test.close()
        except Exception as e:
            logging.info("Excel creation failed: {}".format(e))
            sys.exit()
    # Add the sheet
    wb_test = xw.Book(excel_workbook_testing)
    for i in ["Exchange"]:
        try:
            wb_test.sheets(i)
        except:
            wb_test.sheets.add(i)
    ex = wb_test.sheets("Exchange")
    ex.range("a:j").value = None
    #Without filter condition. This will pull all the instruments.
    ex.range("a1").value = get_instrument_list(kite)

    #With filter condition. This will pull last 2 months of 2023 expiry in NIFTY , BANKNIFTY and FINNIFTY options.
    # ex.range("a1").value = get_instrument_list(kite,"NFO-OPT", 2, [2023], ['NIFTY', 'BANKNIFTY', 'FINNIFTY'])
