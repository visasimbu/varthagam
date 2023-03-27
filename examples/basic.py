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
        # commented for temp simbu
        # self.ex.range("a1").value = get_instrument_list(kite, "NFO-OPT", 2, [2023], ['NIFTY', 'BANKNIFTY', 'FINNIFTY'])
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
             "quantity", "overnight_quantity", "multiplier", "average_price",
             "close_price", "last_price", "value", "pnl", "m2m", "unrealised",
             "realised", "buy_quantity", "buy_price", "buy_value",
             "buy_m2m", "sell_quantity", "sell_price", "sell_value", "sell_m2m",
             "day_buy_quantity", "day_buy_price", "day_buy_value", "day_sell_quantity",
             "day_sell_price", "day_sell_value"]
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
if users_obj is None:
    users_obj = []


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
        current_user = users_obj[i]
        logging.info("Get the position for :" + current_user.userid)
        try:
            excel_obj.wb.sheets(current_user.userid + '_positions')
            excel_obj.wb.sheets(users_obj[i].userid + '_positionsChart')
        except:
            excel_obj.wb.sheets.add(current_user.userid + '_positions')
            excel_obj.wb.sheets.add(current_user.userid + '_positionsChart')

        try:

            current_user.positions = get_positions(current_user.kite)

            now = datetime.datetime.now()
            print(now)
            # datetime.datetime(2009, 1, 6, 15, 8, 24, 78915)
            # print(now)

            if (len(current_user.positions.index)) != 0:
                all_postions_mtm = current_user.positions['M2M'].sum()
                all_postions_pnl = current_user.positions['Pnl'].sum()
                current_user.positions.loc["Total", "M2M"] = all_postions_mtm
                current_user.positions.loc["Total", "Pnl"] = all_postions_pnl
                excel_obj.wb.sheets(current_user.userid + "_positions").range("a1").value = current_user.positions
                if (len(current_user.sum_positions.index)) == 0:
                    logging.info('sum positions is none')
                    current_user.sum_positions = pd.DataFrame(
                        {"M2M": [all_postions_mtm], "Pnl": [all_postions_pnl], "Time": [now]})
                else:
                    logging.info('sum positions is NOT none')
                    current_user.sum_positions.loc[len(current_user.sum_positions.index)] = [all_postions_mtm,
                                                                                             all_postions_pnl, now]
                excel_obj.wb.sheets(current_user.userid + '_positionsChart').range(
                    "a1").value = current_user.sum_positions
            else:
                logging.info('There is no positions for the given user :' + current_user.userid)

            # rng = excel_obj.wb.sheets(current_user.userid + '_positionsChart').cells.last_cell
            # print(rng)
            # excel_obj.wb.sheets(current_user.userid + "_positionsChart").range(rng).value = current_user.sum_positions

        except Exception as e:
            logging.info("Error on get positions: {}".format(e))


if __name__ == '__main__':
    if (len(users_obj)) == 0:
        logging.info("No valid user available. Update the credentials in tokens/input.json file ")
        print("No valid user available. Update the credentials in tokens/input.json file")
    else:
        load_orderbook()
    while True:
        if (len(users_obj)) == 0:
            logging.info("No valid user available. Update the credentials in tokens/input.json file ")
            print("No valid user available. Update the credentials in tokens/input.json file")
            break
        load_positions()
        time.sleep(5)
