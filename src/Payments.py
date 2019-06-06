# -*- coding: utf-8 -*-
"""
Created on Wed May 15 22:11:05 2019

@author: v.shkaberda
"""

from db_connect import DBConnect
from log_error import writelog
from pyodbc import Error as SQLError
import sys
import tkPayments as tkp


def main():
    # Check connection to db and permission to work with app
    try:
        with DBConnect(server='s-kv-center-s59', db='LogisticFinance') as sql:
            access_permitted = sql.access_check()
            if not access_permitted:
                tkp.AccessError()
                sys.exit()

            # load references
            user_info = sql.get_user_info()
            refs = {'connection': sql,
                    'userID': user_info[0],
                    'shortusername': user_info[1],
                    'mvz': sql.get_MVZ()
                    }

            # Run app
            app = tkp.PaymentApp(**refs)
            app.mainloop()


    except SQLError as e:
        #writelog(e)
        # login failed
        if e.args[0] == '42000':
            tkp.LoginError()
        else:
            print(e)  # what to do?


if __name__ == '__main__':
    main()
    input('Press Enter...')
