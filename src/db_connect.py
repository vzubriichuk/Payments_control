# -*- coding: utf-8 -*-
"""
Created on Thu Dec 20 09:43:48 2018

@author: v.shkaberda
"""
import pyodbc


class DBConnect(object):
    ''' Provides connection to database and functions to work with server.
    '''
    def __init__(self, *, server, db):
        self._server = server
        self._db = db

    def __enter__(self):
        # Connection properties
        conn_str = (
            'Driver={{SQL Server}};'
            'Server={0};'
            'Database={1};'
            'Trusted_Connection=yes;'.format(self._server, self._db)
            )
        self.__db = pyodbc.connect(conn_str)
        self.__cursor = self.__db.cursor()
        return self

    def __exit__(self, type, value, traceback):
        self.__db.close()

    def access_check(self):
        ''' Check user permission.
            If access prmitted returns True, otherwise None.
        '''
        self.__cursor.execute('exec [payment].[Access_Check]')
        access = self.__cursor.fetchone()
        # return access[0] if access else 0
        if access and access[0] == 1:
            return True

    def get_user_info(self):
        self.__cursor.execute("select UserID, ShortUserName \
                              from payment.People \
                              where UserLogin = right(ORIGINAL_LOGIN(), charindex('\\', ORIGINAL_LOGIN()))")
        return self.__cursor.fetchone()

    def get_approvelist(self, userID):
        query = '''
        select ID, ShortUserName, cast(date_created as smalldatetime) as date_created,
           MVZ, OfficeID, ContragentID, date_planed,
           SumNoTax, cast(SumNoTax * ((100 + Tax) / 100.0) as numeric(9, 2)),
           p.ValueName as StatusName, pl.Description
        from payment.PaymentsList pl
        join payment.People pp on pl.UserID = pp.UserID
        join dbo.GlobalParamsLines p on pl.StatusID = p.idParamsLines
                                    and p.idParams = 2
                                    and p.Enabled = 1
                                    and pl.StatusID = 1
            '''
        self.__cursor.execute(query)
        res = self.__cursor.fetchall()
        return res

    def get_discardlist(self, userID):
        query = '''
        select ID, cast(date_created as date) as date_created,
               MVZ, OfficeID, date_planed, SumNoTax,
               LEFT(Description, 50) + IIF(LEN(Description) > 50, ' ...', '') as ShortDesc
        from payment.PaymentsList
        where StatusID = 1 and UserID = ?
        '''
        self.__cursor.execute(query, userID)
        return self.__cursor.fetchall()

    def get_MVZ(self):
#        self.__cursor.execute("select distinct MVZ_Name \
#                                from LogisticFinance.BTool.aid_Access_Matrix_Actual_Detail \
#                                where UserID = 2")
        self.__cursor.execute("select FullName, SAPmvz \
                                from LogisticFinance.BTool.aid_CostObject_Detail \
                                where SAPmvz != 'пусто'")
        res = self.__cursor.fetchall()
        return res

    def get_paymentslist(self):
        self.__cursor.execute('''
        select ID, ShortUserName, cast(date_created as smalldatetime) as date_created,
           MVZ, OfficeID, ContragentID, date_planed,
           SumNoTax, cast(SumNoTax * ((100 + Tax) / 100.0) as numeric(9, 2)),
           p.ValueName as StatusName, pl.Description
        from payment.PaymentsList pl
        join payment.People pp on pl.UserID = pp.UserID
        join dbo.GlobalParamsLines p on pl.StatusID = p.idParamsLines
                                    and p.idParams = 2
                                    and p.Enabled = 1
            ''')
        res = self.__cursor.fetchall()
        return res

    def raw_query(self, query):
        ''' Takes the query and returns output from db.
        '''
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    def update_confirmed(self, confirmID, is_confirmed):
        query = '''
        UPDATE payment.PaymentsList
        SET StatusID = ?
        where ID = ?
        '''
        self.__cursor.execute(query, (3 if is_confirmed else 4, confirmID))
        self.__db.commit()

    def update_discarded(self, discardID):
        query = '''
        UPDATE payment.PaymentsList
        SET StatusID = 2
        where ID = ?
        '''
        discardID = [tuple(str(dID)) for dID in discardID]
        self.__cursor.executemany(query, discardID)
        self.__db.commit()


if __name__ == '__main__':
    with DBConnect(server='s-kv-center-s64', db='CB') as sql:
        query = 'select 42'
        assert sql.raw_query(query)[0][0] == 42, 'Server returns no output.'
    print('Connected successfully.')
    input('Press Enter to exit...')
