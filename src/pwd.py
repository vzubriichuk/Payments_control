#!/usr/bin/env python
# coding:utf-8
"""
Author : Vitaliy Zubriichuk
Contact : v@zubr.kiev.ua
Time    : 08.04.2020 19:04
"""
def access_return(n):
    conn_str = {
        "Driver": '{SQL Server}',
        "Server": 's-kv-center-s59',
        "DB": None,
        "UID": None,
        "PWD": None
    }
    if n == 1:
        conn_str.update(DB='AnalyticReports')
        conn_str.update(UID='j-PaymentDev')
        conn_str.update(PWD='V41m9i1Z90XPpdfaPICy')
    else:
        conn_str.update(DB='LogisticFinance')
        conn_str.update(UID='j-LogisticFinance-Spec')
        conn_str.update(PWD='dfi6qdVUI3BhJ8kJtH0i')
    return conn_str

