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


def get_live_data(instruments):
    global kite, live_data
    try:
        live_data
        print("50")
    except:
        print("51")
        live_data = {}
    try:
        print("52")
        live_data = kite.quote(instruments)
    except Exception as e:
        print("53")
        print(f"Get live data failed..{{{e}}}")
        pass
    return live_data
