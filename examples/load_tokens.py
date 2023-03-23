import json
import logging
import pandas as pd
from kiteconnect import KiteConnect

logging.basicConfig(filename='../log/loadtoken.log', level=logging.DEBUG)
log = logging.getLogger(__name__)


def load_all_tokens():
    try:
        with open('../tokens/input.json', 'r') as json_file:
            data = json.load(json_file)
        obj_holder = [LoadAllUsers(data["user_lists"][each_user_index]["nickname"],
                                   data["user_lists"][each_user_index]["userid"],
                                   data["user_lists"][each_user_index]["enctoken"]
                                   ) for each_user_index in
                      range(len(data["user_lists"]))]

        for json_data_index in range(len(data["user_lists"])):
            obj_holder[json_data_index].validate_users_credentials()

        valid_user_obj = []
        Not_valid_user_obj = []
        for index in range(len(obj_holder)):
            if obj_holder[index].is_valid_user:
                valid_user_obj.append(obj_holder[index])
            else:
                Not_valid_user_obj.append(obj_holder[index])

        logging.info("******************************")
        logging.info("List of In Valid users :  ")
        for i in range(len(Not_valid_user_obj)):
            logging.info(Not_valid_user_obj[i].userid)
        logging.info("******************************")
        logging.info("==========================")
        logging.info("List of Valid users :  ")
        for i in range(len(valid_user_obj)):
            logging.info(valid_user_obj[i].userid)
        logging.info("==========================")
        return valid_user_obj
    except Exception as e:
        logging.info("Error on load_all_tokens function: {}".format(e))


class LoadAllUsers:
    def __init__(self, nickname, userid, enctoken):
        self.nickname = nickname
        self.userid = userid
        self.enctoken = enctoken
        self.is_valid_user = False
        self.positions = pd.DataFrame({})
        self.sum_positions = pd.DataFrame({})
        self.kite = KiteConnect(enc_token=enctoken)

    def validate_users_credentials(self):
        logging.info("Validating user " + self.nickname)
        try:
            if self.kite.profile()["user_id"] != self.userid:
                logging.info("In valid enctoken for the user ID:" + self.userid)
                self.is_valid_user = False
            else:
                logging.info("User ID:" + self.userid + " validated successfully ")
                self.is_valid_user = True
        except Exception as e:
            logging.info("Error while we perform validating user : " + self.nickname)
            logging.info("Error: {}".format(e))


if __name__ == '__main__':
    load_all_tokens()
