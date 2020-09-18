#!/usr/bin/env python
# coding:utf-8
"""
Author : Vitaliy Zubriichuk
Contact : v@zubr.kiev.ua
Time    : 08.04.2020 19:04
"""
def access_return(n):
    conn_info = {
        "Driver": '{SQL Server}',
        "Server": 's-kv-center-s59',
        "DB": None,
        "UID": None,
        "PWD": None
    }
    if n == 1:
        conn_info.update(DB='AnalyticReports')
        conn_info.update(UID='j-PaymentDev')
        conn_info.update(PWD='V41m9i1Z90XPpdfaPICy')
    else:
        conn_info.update(DB='LogisticFinance')
        conn_info.update(UID='j-LogisticFinance-Spec')
        conn_info.update(PWD='dfi6qdVUI3BhJ8kJtH0i')
    return conn_info

