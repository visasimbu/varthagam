import datetime
import logging
import os
import sys
import time
import pandas as pd
import xlwings as xw

from kiteconnect import KiteConnect
from getinstruments import get_instrument_list
from load_tokens import load_all_tokens

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

class ExcelTrade:
    def __init__(self,
                 debug=False):

        self.orders = None
        self.debug = debug

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
        self.dt.range(f"a1:q1").value = ["Sr No", "Symbol", "Open", "High", "Low", "LTP", "volume", "Vwap",
                                         "Best Bid Price",
                                         "Best Ask Price", "Close", "Qty", "Direction", "Entry Signal", "Exit Signal",
                                         "Entry",
                                         "Exit"]
        # For all the instruments.
        # ex.range("a1").value = get_instrument_list()
        self.ex.range("a1").value = get_instrument_list(kite, "NFO-OPT", 2, [2023], ['NIFTY', 'BANKNIFTY', 'FINNIFTY'])
        logging.info("Excel started  !!! Start your trading")


def get_orderbook(kite):
    # global orders
    try:
        orders = None
    except:
        orders = {}
    try:
        print('In getting order book')
        data = pd.DataFrame(kite.orders())
        # print(data.info())
        # print(data.describe())
        # data = data[data["tag"] == "FromPython"]
        data = data.filter(
            ["order_timestamp", "exchange", "tradingsymbol", "transaction_type", "quantity", "average_price",
             "trigger_price", "filled_quantity", "pending_quantity", "cancelled_quantity",
             "market_protection",
             # "meta",
             "tag", "guid",
             "status",
             "status_message_raw"]
        )

        data.columns = data.columns.str.replace("_", " ")
        data.columns = data.columns.str.title()
        # data = data.set_index(["Order Timestamp"], drop=True)
        data = data.sort_index(ascending=True)
        orders = data
    except Exception as e:
        print(e)
        pass
    return orders


def get_positions(kite):
    # global orders
    try:
        positions = None
    except:
        positions = {}
    try:
        print('In getting positions')
        data = pd.DataFrame(kite.positions()["net"])
        # print(data.describe())
        # print(data.info())
        # data = data[data["tag"] == "FromPython"]
        data = data.filter(
            ["tradingsymbol",
             "quantity","overnight_quantity","multiplier","average_price",
             "close_price","last_price","value","pnl","m2m","unrealised",
             "realised","buy_quantity","buy_price","buy_value",
             "buy_m2m","sell_quantity","sell_price","sell_value","sell_m2m",
             "day_buy_quantity","day_buy_price","day_buy_value","day_sell_quantity",
             "day_sell_price","day_sell_value"]
        )

        data.columns = data.columns.str.replace("_", " ")
        data.columns = data.columns.str.title()
        # data = data.set_index(["Order Timestamp"], drop=True)
        data = data.sort_index(ascending=True)
        positions = data
    except Exception as e:
        logging.info("Excel creation failed: {}".format(e))
        pass
    return positions


#
# def start_excel():
#     global kite, live_data
#     logging.info("Excel starting...")
#     # Create Excel file if not available.
#     if not os.path.exists(excel_workbook):
#         try:
#             wb = xw.Book()
#             wb.save(excel_workbook)
#             wb.close()
#         except Exception as e:
#             logging.info("Excel creation failed: {}".format(e))
#             sys.exit()
#     # Add the sheet
#     wb = xw.Book(excel_workbook)
#     for i in ["Data", "Exchange", "orderBook"]:
#         try:
#             wb.sheets(i)
#         except:
#             wb.sheets.add(i)
#     # Initialize sheet Variable
#     dt = wb.sheets("Data")
#     ex = wb.sheets("Exchange")
#     ob = wb.sheets("orderBook")
#
#     # Clean the data in all 3 sheets in Excel file
#     ex.range("a:j").value = ob.range("a:h").value = dt.range("p:q").value = None
#     # Initialize data sheet with column name
#     dt.range(f"a1:q1").value = ["Sr No", "Symbol", "Open", "High", "Low", "LTP", "volume", "Vwap", "Best Bid Price",
#                                 "Best Ask Price", "Close", "Qty", "Direction", "Entry Signal", "Exit Signal", "Entry",
#                                 "Exit"]
#     # For all the instruments.
#     # ex.range("a1").value = get_instrument_list()
#     ex.range("a1").value = get_instrument_list(kite, "NFO-OPT", 2, [2023], ['NIFTY', 'BANKNIFTY', 'FINNIFTY'])
#     logging.info("Excel started  !!! Start your trading")

excel_obj = ExcelTrade()
users_obj = load_all_tokens()


def load_orderbook():
    for i in range(len(users_obj)):
        logging.info("Get the order book for :" + users_obj[i].userid)
        try:
            excel_obj.wb.sheets(users_obj[i].userid + '_orderbook')
        except:
            excel_obj.wb.sheets.add(users_obj[i].userid + '_orderbook')

        try:
            excel_obj.wb.sheets(users_obj[i].userid + "_orderbook").range("a1").value = get_orderbook(users_obj[i].kite)
        except Exception as e:
            logging.info("Error on get orderbook: {}".format(e))


def load_positions():
    for i in range(len(users_obj)):
        logging.info("Get the position for :" + users_obj[i].userid)
        try:
            excel_obj.wb.sheets(users_obj[i].userid + '_positions')
        except:
            excel_obj.wb.sheets.add(users_obj[i].userid + '_positions')

        try:
            positions = get_positions(users_obj[i].kite)
            # print(positions.info())
            all_postions_mtm = positions['M2M'].sum()
            print(all_postions_mtm)
            positions.loc["Total", "M2M"] = positions.M2M.sum()
            positions.loc["Total", "Pnl"] = positions.Pnl.sum()

            excel_obj.wb.sheets(users_obj[i].userid + "_positions").range("a1").value = positions
        except Exception as e:
            logging.info("Error on get positions: {}".format(e))


if __name__ == '__main__':
    load_orderbook()
    while True:
        load_positions()
        time.sleep(5)
