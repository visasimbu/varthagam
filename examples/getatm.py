import logging
import datetime
import os
import sys
import time
import pandas as pd

from kiteconnect import KiteConnect
from getinstruments import get_instrument_list
from load_tokens import load_all_tokens

logging.basicConfig(filename='../log/getatm.log', level=logging.DEBUG)
log = logging.getLogger(__name__)

# kite = KiteConnect(enc_token=
#                    "xsdfdsfdsfsdfdsfasdaf==")


current_month = 'APR'  # now.strftime("%b")
current_year = '23'  # now.strftime("%y")
nifty_expiry = datetime.date(2023, 4, 27)
banknifty_expiry = datetime.date(2023, 4, 27)
finnifty_expiry = datetime.date(2023, 4, 25)

users_obj = load_all_tokens()
if users_obj is None:
    users_obj = []
else:
    kite = users_obj[1].kite

today = datetime.datetime.today()
base_path = 'F:/SIMBU/Python/Trading/Zerodha/'
df_atm_straddle = pd.DataFrame({'timestamp': [], 'combined_premium': [], 'atm_strike': [], 'instrument': []})
atm_straddle_workbook = base_path + '_atm_straddle.csv'


class feed_atm:
    def __init__(self, instrument, debug=False):
        self.debug = debug  # True
        self.atm_straddle_workbook = base_path + str(today.strftime("%d-%b-%Y")) + '_' + str(
            instrument) + '_atm_straddle.csv'

        if instrument == 'NIFTY':
            self.expiry = nifty_expiry
        elif instrument == 'BANKNIFTY':
            self.expiry = banknifty_expiry
        elif instrument == 'FINNIFTY':
            self.expiry = finnifty_expiry

        self.create_file(instrument)

    def getATM(self, name, instrument_list_df=None):
        atmstk = 0
        try:
            # tickerid = name + "23APRFUT"
            tickerid = name + current_year + current_month + "FUT"
            tickerid = tickerid.upper()
            ltp_dict = kite.ltp('NFO:' + str(tickerid))
            if bool(ltp_dict):
                ltp = round(ltp_dict.get('NFO:' + str(tickerid), {}).get('last_price'), 0)
            else:
                logging.info(
                    'Error: Receiving null dict on LTP GET call. The given future instrument index is not valid: ' + str(
                        tickerid))
            if instrument_list_df is None:
                logging.debug("instrument list is none. So get the fresh instrument list")
                instrument_df = get_instrument_list(kite, "NFO-OPT", 2, [2023], ['NIFTY', 'BANKNIFTY', 'FINNIFTY'])
            else:
                logging.debug("instrument list is already available")
                instrument_df = instrument_list_df

            df = instrument_df
            # expiry = datetime.date(2023, 3, 29)

            atmdf = df.loc[(df['Name'] == name) & (df['Expiry'] == self.expiry)]
            lst = atmdf['Strike'].tolist()
            atmstk = round(self.closest(lst, ltp))
        except Exception as e:
            logging.info('Error on getting ATM strike : {}'.format(e))
        return atmstk

    def closest(self, lst, K):
        return lst[min(range(len(lst)), key=lambda i: abs(lst[i] - K))]

    def combined_premium(self, instrument, strike):
        instrument_name = 'NFO:' + instrument + current_year + current_month + str(strike)
        instrument_name = instrument_name.upper()
        try:
            ce_ltp = round((kite.ltp(instrument_name + 'CE')).get(instrument_name + 'CE', {}).get('last_price'), 0)
            pe_ltp = round((kite.ltp(instrument_name + 'PE')).get(instrument_name + 'PE', {}).get('last_price'), 0)
        except Exception as e:
            logging.info('Error on Getting LTP of CE and PE:{}'.format(e))
        combined_ltp = ce_ltp + pe_ltp
        return combined_ltp

    def create_file(self, instrument):
        # global atm_straddle_workbook
        # atm_straddle_workbook = base_path + str(today.strftime("%d-%b-%Y")) + '_' + str(
        #     instrument) + '_atm_straddle.csv'
        if not os.path.exists(self.atm_straddle_workbook):
            try:
                df_atm_straddle.to_csv(self.atm_straddle_workbook, mode='w', index=False, header=True)
            except Exception as e:
                logging.info("atm_straddle_workbook csv creation failed: {}".format(e))
                sys.exit()

    def atm_feed_to_file(self, instrument):
        global df_atm_straddle
        now = datetime.datetime.now()
        cp = 0
        atm = self.getATM(instrument, instrument_df)
        if atm != 0:
            cp = self.combined_premium(instrument, atm)
        else:
            logging.info('Error: The given instrument is not valid. Unable to fetch ATM strike')
            # break
        if df_atm_straddle is None:
            df_atm_straddle = pd.DataFrame(
                {'timestamp': [], 'combined_premium': [], 'atm_strike': [], 'instrument': []})
            index_df = 0
        else:
            index_df = len(df_atm_straddle)
        df_atm_straddle.loc[
            index_df, ['timestamp', 'combined_premium', 'atm_strike', 'instrument']] = now, cp, atm, instrument
        df_atm_straddle.to_csv(self.atm_straddle_workbook, mode='a', index=False, header=False)
        df_atm_straddle = None


#
# getATM("NIFTY") or getATM("NIFTY", instrument_df)
# Output: 17000.0...

if __name__ == '__main__':
    nifty_obj = feed_atm('NIFTY')
    banknifty_obj = feed_atm('BANKNIFTY')
    finnifty_obj = feed_atm('FINNIFTY')

    instrument_df = get_instrument_list(kite, "NFO-OPT", 2, [2023], ['NIFTY', 'BANKNIFTY', 'FINNIFTY'])
    logging.info('Getting atm combined premium into files...')
    print('Getting atm combined premium into files...')
    while True:
        nifty_obj.atm_feed_to_file('NIFTY')
        banknifty_obj.atm_feed_to_file('BANKNIFTY')
        finnifty_obj.atm_feed_to_file('FINNIFTY')
        time.sleep(5)
