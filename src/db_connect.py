# -*- coding: utf-8 -*-
"""
Created on Thu Dec 20 09:43:48 2018

@author: v.shkaberda
"""
from functools import wraps
from tkPayments import NetworkError
import pyodbc

def monitor_network_state(method):
    """ Show error message in case of network error
    """
    @wraps(method)
    def wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        except pyodbc.Error as e:
            # Network error
            if e.args[0] in ('01000', '08S01', '08001'):
                NetworkError()
    return wrapper


class DBConnect(object):
    """ Provides connection to database and functions to work with server.
    """
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

    @monitor_network_state
    def access_check(self):
        """ Check user permission.
            If access prmitted returns True, otherwise None.
        """
        self.__cursor.execute('exec [payment].[Access_Check]')
        access = self.__cursor.fetchone()
        # check AccessType and isSuperUser
        if access and (access[0] in (1, 2) or access[1]):
            return True

    @monitor_network_state
    def create_request(self, userID, mvz, office, contragent, csp, plan_date,
                       sumtotal, nds, text):
        query = '''
        exec payment.create_request @UserID = ?,
                                    @MVZ = ?,
                                    @Office = ?,
                                    @Contragent = ?,
                                    @date_planed = ?,
                                    @Description = ?,
                                    @SumNoTax = ?,
                                    @Tax = ?,
                                    @CSP = ?
            '''
        try:
            self.__cursor.execute(query, userID, mvz, office, contragent,
                                  plan_date, text, sumtotal, nds, csp)
            request_allowed = self.__cursor.fetchone()[0]
            self.__db.commit()
            return request_allowed
        except pyodbc.ProgrammingError:
            return

    @monitor_network_state
    def get_user_info(self):
        self.__cursor.execute("select UserID, ShortUserName, AccessType, isSuperUser \
          from payment.People \
          where UserLogin = right(ORIGINAL_LOGIN(), len(ORIGINAL_LOGIN()) - charindex( '\\' , ORIGINAL_LOGIN()))")
        return self.__cursor.fetchone()

    @monitor_network_state
    def get_allowed_initiators(self, UserID, AccessType, isSuperUser):
        query = '''
        exec payment.get_allowed_initiators @UserID = ?,
                                            @AccessType = ?,
                                            @isSuperUser = ?
        '''
        self.__cursor.execute(query, UserID, AccessType, isSuperUser)
        return [(None, 'Все'),] + self.__cursor.fetchall()

    @monitor_network_state
    def get_approvals(self, paymentID):
        query = '''
        exec payment.get_approvals @paymentID = ?
        '''
        self.__cursor.execute(query, paymentID)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_limits_info(self):
        query = '''
        select UserID, UserName, userCreateRequestLimit, resetCreateRequestLimit
        from payment.People
        where AccessType in (1, 2)
            or isSuperUser = 1
        order by UserName
            '''
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_MVZ(self, user_info):
        if user_info.isSuperUser:
            query = '''
            select obj.MVZsap, co.FullName, obj.ServiceName
            from payment.ObjectsList obj
                join BTool.aid_CostObject_Detail co on co.SAPMVZ = obj.MVZsap\n
            '''
        else:
            query = '''
            select obj.MVZsap, co.FullName, obj.ServiceName
            from payment.ObjectsList obj
                join BTool.aid_CostObject_Detail co on co.SAPMVZ = obj.MVZsap
                join payment.User_Approvals_Ref appr on obj.ID = appr.ObjectID
            where appr.UserID = {}\n
            '''.format(user_info.UserID)
            if user_info.AccessType == 1:
                query += ("and cast(getdate() as date) between appr.activeFrom"
                          " and isnull(appr.activeTo, '20990101')\n")
        query += "order by co.FullName"
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_paymentslist(self, user_info, initiator, mvz, office,
                         plan_date_m, plan_date_y, sumtotal_from, sumtotal_to,
                         nds, just_for_approval):
        """ Generate query according to user's acces type and filters.
        """
        query = '''
        select pl.ID as ID, pl.UserID as InitiatorID,
           'Лог-' + replace(convert(varchar, date_created, 102),'.','') + '_' + cast(pl.ID as varchar(7)) as Num,
           pp.ShortUserName, cast(date_created as date) as date_created,
           cast(date_created as smalldatetime) as datetime_created,
           isnull(CSP, '') as CSP, obj.MVZsap, co.FullName, obj.ServiceName,
           isnull(Contragent, '') as Contragent, date_planed,
           SumNoTax, cast(SumNoTax * ((100 + Tax) / 100.0) as numeric(11, 2)),
           p.ValueName as StatusName, p.ValueDescription, pl.Description, appr.UserID,
           case when pl.StatusID = 1 then isnull(pappr.ShortUserName, '') else '' end as approval
        from payment.PaymentsList pl
        join payment.ObjectsList obj on pl.ObjectID = obj.ID
        join BTool.aid_CostObject_Detail co on co.SAPMVZ = obj.MVZsap
        join payment.People pp on pl.UserID = pp.UserID
        join dbo.GlobalParamsLines p on pl.StatusID = p.idParamsLines
                                    and p.idParams = 2
                                    and p.Enabled = 1
        -- approvals
        left join payment.PaymentsApproval appr on pl.ID = appr.PaymentID
                                           and appr.is_active_approval = 1
        left join payment.People pappr on appr.UserID = pappr.UserID
        where 1=1
        '''
        if just_for_approval:
            query += 'and pl.StatusID = 1 and appr.UserID = {}\n'.format(user_info.UserID)
        else:
            if not user_info.isSuperUser:
                query += "and (pl.UserID = {0} or exists(select * from payment.PaymentsApproval _appr \
                        where pl.ID = _appr.PaymentID and _appr.UserID = {0}))\n".format(user_info.UserID)
            if initiator:
                query += "and pl.UserID = {}\n".format(initiator)
            if mvz:
                query += "and obj.MVZsap = '{}'\n".format(mvz)
            if office:
                query += "and obj.ServiceName = '{}'\n".format(office)
            if plan_date_y:
                query += "and year(date_planed) = {}\n".format(plan_date_y)
            if plan_date_m:
                query += "and month(date_planed) in ({})\n".format(plan_date_m)
            if sumtotal_from:
                query += "and SumNoTax >= {}\n".format(sumtotal_from)
            if sumtotal_to:
                query += "and SumNoTax <= {}\n".format(sumtotal_to)
            if not nds == -1:
                query += "and Tax = {}\n".format(nds)
        query += "order by ID DESC" # the same as created(datetime) DESC
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def raw_query(self, query):
        """ Takes the query and returns output from db.
        """
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def update_confirmed(self, userID, paymentID, is_approved):
        query = '''
        exec payment.approve_request @UserID = ?,
                                     @paymentID = ?,
                                     @is_approved = ?
        '''
        self.__cursor.execute(query, userID, paymentID, is_approved)
        self.__db.commit()

    @monitor_network_state
    def update_discarded(self, discardID):
        query = '''
        UPDATE payment.PaymentsList
        SET StatusID = 2
        where ID = ?
        '''
        self.__cursor.execute(query, discardID)
        self.__db.commit()

    @monitor_network_state
    def update_limits(self, limits):
        query = '''
        UPDATE payment.People
        SET resetCreateRequestLimit = ?, userCreateRequestLimit = ?
        where UserID = ?
        '''
        limits = [tuple(reversed(info)) for info in limits]
        try:
            self.__cursor.executemany(query, limits)
            self.__db.commit()
            return 1
        except pyodbc.ProgrammingError:
            return 0


if __name__ == '__main__':
    with DBConnect(server='s-kv-center-s64', db='CB') as sql:
        query = 'select 42'
        assert sql.raw_query(query)[0][0] == 42, 'Server returns no output.'
    print('Connected successfully.')
    input('Press Enter to exit...')
