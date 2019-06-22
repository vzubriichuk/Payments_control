# -*- coding: utf-8 -*-
"""
Created on Thu Dec 20 09:43:48 2018

@author: v.shkaberda
"""
import pyodbc


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

    def access_check(self):
        """ Check user permission.
            If access prmitted returns True, otherwise None.
        """
        self.__cursor.execute('exec [payment].[Access_Check]')
        access = self.__cursor.fetchone()
        # check AccessType and isSuperUser
        if access and (access[0] in (1, 2) or access[1]):
            return True

    def create_request(self, userID, mvz, office, contragent, plan_date,
                       sumtotal, nds, text):
        query = '''
        exec payment.create_request @UserID = ?,
                                    @MVZ = ?,
                                    @OfficeID = ?,
                                    @ContragentID = ?,
                                    @date_planed = ?,
                                    @Description = ?,
                                    @SumNoTax = ?,
                                    @Tax = ?
            '''
        try:
            self.__cursor.execute(query, userID, mvz, office, contragent, plan_date,
                                  text, sumtotal, nds)
            self.__db.commit()
            return 1
        except pyodbc.ProgrammingError:
            return 0

    def get_user_info(self):
        self.__cursor.execute("select UserID, ShortUserName, AccessType, isSuperUser \
          from payment.People \
          where UserLogin = right(ORIGINAL_LOGIN(), len(ORIGINAL_LOGIN()) - charindex( '\\' , ORIGINAL_LOGIN()))")
        return self.__cursor.fetchone()

    def get_allowed_initiators(self, UserID, AccessType, isSuperUser):
        if isSuperUser:
            query = '''
            select UserID, UserName
            from payment.People
            where AccessType in (1, 2, 3)
            order by UserName
            '''
        elif AccessType == 1:
            query = '''
            select UserID, UserName
            from payment.People
            where UserID = {}
            '''.format(UserID)
        elif AccessType == 2:
            query = '''
            select UserID, UserName
            from payment.People
            where AccessType in (1, 2, 3)
                or UserID = {}
            order by UserName
            '''.format(UserID)
        self.__cursor.execute(query)
        return [(None, 'Все'),] + self.__cursor.fetchall()

    def get_approvals(self, paymentID):
        query = '''
        select pappr.ShortUserName as approval,
        case appr.is_approved when 0 then 'Отклонил(-а) ' + CONVERT(nvarchar, date_modified, 20)
                              when 1 then 'Утвердил(-а) ' + CONVERT(nvarchar, date_modified, 20)
                              else '' end as status
        from payment.PaymentsApproval appr
            join payment.People pappr on appr.UserID = pappr.UserID
        where appr.PaymentID = ?
        order by approval_order ASC
            '''
        self.__cursor.execute(query, paymentID)
        return self.__cursor.fetchall()

    def get_discardlist(self, userID):
        query = '''
        select ID as ID,
           'Лог-' + replace(convert(varchar, date_created, 102),'.','') + '_' + cast(ID as varchar(7)) as Num,
           cast(date_created as date) as date_created, '' as CSP,
           MVZ, OfficeID, date_planed, SumNoTax,
           LEFT(Description, 50) + IIF(LEN(Description) > 50, ' ...', '') as ShortDesc
        from payment.PaymentsList
        where StatusID = 1 and UserID = ?
        '''
        self.__cursor.execute(query, userID)
        return self.__cursor.fetchall()

    def get_MVZ(self):
        self.__cursor.execute("select FullName, SAPmvz \
                               from LogisticFinance.BTool.aid_CostObject_Detail \
                               where SAPmvz != 'пусто' \
                               order by FullName")
        return self.__cursor.fetchall()

    def get_OKPO(self):
        self.__cursor.execute("select top 20 pac.ValueProperty as OKPO, a.name \
                    from CB.dbo.PropertyAdressChar pac \
                        join CB.dbo.Adress a with(nolock) on pac.AdressID = a.ID \
                    where PropertyID = 2")
        return self.__cursor.fetchall()

    def get_paymentslist(self, user_info, initiator, mvz, office, contragent,
                         plan_date_m, plan_date_y, sumtotal_from, sumtotal_to,
                         nds, just_for_approval):
        """ Generate query according to user's acces type and filters.
        """
        query = '''
        select pl.ID as ID, pl.UserID as InitiatorID,
           'Лог-' + replace(convert(varchar, date_created, 102),'.','') + '_' + cast(pl.ID as varchar(7)) as Num,
           pp.ShortUserName, cast(date_created as date) as date_created,
           cast(date_created as smalldatetime) as datetime_created, '' as CSP,
           MVZ, OfficeID, isnull(ContragentID, '') as ContragentID, date_planed,
           SumNoTax, cast(SumNoTax * ((100 + Tax) / 100.0) as numeric(11, 2)),
           p.ValueName as StatusName, p.ValueDescription, pl.Description,
           case when pl.StatusID = 1 then isnull(pappr.ShortUserName, '') else '' end as approval
        from payment.PaymentsList pl
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
                query += 'and (pl.UserID = {0} or exists(select * from payment.PaymentsApproval _appr \
                        where pl.ID = _appr.PaymentID and _appr.UserID = {0}))\n'.format(user_info.UserID)
            if initiator:
                query += 'and pl.UserID = {}\n'.format(initiator)
            if mvz:
                query += 'and MVZ = {}\n'.format(mvz)
            if office:
                query += 'and OfficeID = {}\n'.format(office)
            if contragent:
                query += 'and ContragentID = {}\n'.format(contragent)
            if plan_date_y:
                query += 'and year(date_planed) = {}\n'.format(plan_date_y)
            if plan_date_m:
                query += 'and month(date_planed) = {}\n'.format(plan_date_m)
            if sumtotal_from:
                query += 'and SumNoTax >= {}\n'.format(sumtotal_from)
            if sumtotal_to:
                query += 'and SumNoTax <= {}\n'.format(sumtotal_to)
            if not nds == -1:
                query += 'and Tax = {}\n'.format(nds)
        query += 'order by ID DESC' # the same as created(datetime) DESC
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    def raw_query(self, query):
        """ Takes the query and returns output from db.
        """
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    def update_confirmed(self, userID, paymentID, is_approved):
        query = '''
        exec payment.approve_request @UserID = ?,
                                     @paymentID = ?,
                                     @is_approved = ?
        '''
        self.__cursor.execute(query, userID, paymentID, is_approved)
        self.__db.commit()

    def update_discarded(self, discardID):
        query = '''
        UPDATE payment.PaymentsList
        SET StatusID = 2
        where ID = ?
        '''
        self.__cursor.execute(query, discardID)
        self.__db.commit()


if __name__ == '__main__':
    with DBConnect(server='s-kv-center-s64', db='CB') as sql:
        query = 'select 42'
        assert sql.raw_query(query)[0][0] == 42, 'Server returns no output.'
    print('Connected successfully.')
    input('Press Enter to exit...')
