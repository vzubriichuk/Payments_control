# -*- coding: utf-8 -*-
"""
Created on Wed May 15 22:11:05 2019

@author: v.shkaberda
"""
from collections import namedtuple
from db_connect import DBConnect
#from log_error import writelog
# Temporary shut down logging to check if hanging of payments_checker.py
# is an issue because of attempt to access the same log file simultaniously
from os import path
from pyodbc import Error as SQLError
from singleinstance import Singleinstance
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

            UserInfo = namedtuple('UserInfo', ['UserID', 'ShortUserName',
                                               'AccessType', 'isSuperUser'])

            # load references
            user_info = UserInfo(*sql.get_user_info())
            # Restriction: users in approvals_for_first_stage
            # should have different names to be distinguished
            refs = {'connection': sql,
                    'user_info': user_info,
                    'mvz': sql.get_MVZ(user_info),
                    'categories': sql.get_categories(user_info),
                    'allowed_initiators': sql.get_allowed_initiators(user_info.UserID,
                                                                     user_info.AccessType,
                                                                     user_info.isSuperUser),
                    'approvals_for_first_stage': sql.get_approvals_for_first_stage(),
                    'status_list': sql.get_status_list()
                    }
            # Run app
            app = tkp.PaymentApp(**refs)
            app.mainloop()

    except SQLError as e:
        # login failed
        if e.args[0] in ('28000', '42000'):
            #writelog(e)
            tkp.LoginError()
        else:
            raise


if __name__ == '__main__':
    try:
        fname = path.basename(__file__)
        myapp = Singleinstance(fname)
        if myapp.aleradyrunning():
            sys.exit()
        main()
    except Exception as e:
        #writelog(e)
        print(e)
    finally:
        sys.exit()
