import datetime
import logging
import os
import sys
import time
import pandas as pd
import xlwings as xw

from kiteconnect import KiteConnect
from getinstruments import get_instrument_list

logging.basicConfig(filename='mylog', level=logging.DEBUG)

kite = KiteConnect(enc_token=
                   "cpTNh3/uYNqEIl0oMwYRSJtphUxXeTRU9WbV2HbA73AYeU2boXpIxiwOhi2S9aihMBotk2uR6CUP9NNCkAGAOycd4RJoRI9lBDp5Hfot5YTEKQYoZLO13Q==")
excel_workbook = 'F:\SIMBU\Python\Trading\Zerodha\TradeOn.xlsx'


# Redirect the user to the login url obtained
# from kite.login_url(), and receive the request_token
# from the registered redirect url after the login flow.
# Once you have the request_token, obtain the access_token
# as follows.

# data = kite.generate_session("request_token_here", secret="your_secret")
# kite.set_access_token(data["access_token"])

# Fetch all orders
# kite.orders()

# Get instruments
# print(kite.instruments('NSE'))

# Place an order
# try:
#     order_id = kite.place_order(
#         variety=kite.VARIETY_AMO,
#         exchange=kite.EXCHANGE_NSE,
#         tradingsymbol="INFY",
#         transaction_type=kite.TRANSACTION_TYPE_BUY,
#         quantity=1,
#         product=kite.PRODUCT_CNC,
#         order_type=kite.ORDER_TYPE_LIMIT,
#         price=1450
#     )
#     logging.info("Order placed. ID is: {}".format(order_id))
# except Exception as e:
#     print('inside exception')
#     logging.info("Order placement failed: {}".format(e))

class ExcelTrade(object):
    def __init__(self,
                 kite,
                 debug=False):

        self.orders = None
        self.debug = debug

        self.kite = kite
        # self.live_date = live_data

        logging.info("Excel starting...")
        # Create Excel file if not available.
        if not os.path.exists(excel_workbook):
            try:
                self.wb = xw.Book()
                self.wb.save(excel_workbook)
                self.wb.close()
            except Exception as e:
                logging.info("Excel creation failed: {}".format(e))
                sys.exit()
        # Add the sheet
        self.wb = xw.Book(excel_workbook)
        for i in ["Data", "Exchange", "orderBook"]:
            try:
                self.wb.sheets(i)
            except:
                self.wb.sheets.add(i)
        # Initialize sheet Variable
        self.dt = self.wb.sheets("Data")
        self.ex = self.wb.sheets("Exchange")
        self.ob = self.wb.sheets("orderBook")
        # Clean the data in all 3 sheets in Excel file
        self.ex.range("a:j").value = self.ob.range("a:h").value = self.dt.range("p:q").value = None
        # Initialize data sheet with column name
        self.dt.range(f"a1:q1").value = ["Sr No", "Symbol", "Open", "High", "Low", "LTP", "volume", "Vwap", "Best Bid Price",
                                    "Best Ask Price", "Close", "Qty", "Direction", "Entry Signal", "Exit Signal",
                                    "Entry",
                                    "Exit"]
        # For all the instruments.
        # ex.range("a1").value = get_instrument_list()
        self.ex.range("a1").value = get_instrument_list(kite, "NFO-OPT", 2, [2023], ['NIFTY', 'BANKNIFTY', 'FINNIFTY'])
        logging.info("Excel started  !!! Start your trading")

    def get_orderbook(self):
        # global orders
        try:
            self.orders
        except:
            self.orders = {}
        try:
            print('In getting order book')
            data = pd.DataFrame(kite.orders())
            print(data.info())
            # data = data[data["tag"] == "FromPython"]
            data = data.filter(
                ["order_timestamp", "exchange", "tradingsymbol", "transaction_type", "quantity", "average_price",
                 "trigger_price","filled_quantity","pending_quantity","cancelled_quantity",
                 "market_protection",
                 # "meta",
                 "tag","guid",
                 "status",
                 "status_message_raw"]
                )


            data.columns = data.columns.str.replace("_", " ")
            data.columns = data.columns.str.title()
            # data = data.set_index(["Order Timestamp"], drop=True)
            data = data.sort_index(ascending=True)
            self.orders = data
        except Exception as e:
            print(e)
            pass
        return self.orders


def start_excel():
    global kite, live_data
    logging.info("Excel starting...")
    # Create Excel file if not available.
    if not os.path.exists(excel_workbook):
        try:
            wb = xw.Book()
            wb.save(excel_workbook)
            wb.close()
        except Exception as e:
            logging.info("Excel creation failed: {}".format(e))
            sys.exit()
    # Add the sheet
    wb = xw.Book(excel_workbook)
    for i in ["Data", "Exchange", "orderBook"]:
        try:
            wb.sheets(i)
        except:
            wb.sheets.add(i)
    # Initialize sheet Variable
    dt = wb.sheets("Data")
    ex = wb.sheets("Exchange")
    ob = wb.sheets("orderBook")

    # Clean the data in all 3 sheets in Excel file
    ex.range("a:j").value = ob.range("a:h").value = dt.range("p:q").value = None
    # Initialize data sheet with column name
    dt.range(f"a1:q1").value = ["Sr No", "Symbol", "Open", "High", "Low", "LTP", "volume", "Vwap", "Best Bid Price",
                                "Best Ask Price", "Close", "Qty", "Direction", "Entry Signal", "Exit Signal", "Entry",
                                "Exit"]
    # For all the instruments.
    # ex.range("a1").value = get_instrument_list()
    ex.range("a1").value = get_instrument_list(kite, "NFO-OPT", 2, [2023], ['NIFTY', 'BANKNIFTY', 'FINNIFTY'])
    logging.info("Excel started  !!! Start your trading")



if __name__ == '__main__':
    # start_excel()
    excelTrade = ExcelTrade(kite)
    excelTrade.ob.range("a1").value = excelTrade.get_orderbook()
