# -*- coding: utf-8 -*-
# @Author  : chenyihui
# @Email   : chenyihui10@126.com
# @Time    :2019-2-21 20:04
# WARNING! All changes made in this file will be lost!
import pandas as pd
import os
from datetime import datetime
from sqlalchemy import create_engine
# Path = os.path.abspath('..')
# print(Path)
from config import MysqlConfig
# tdate = pd.read_excel(r'E:\myprog\tradedate.xlsx')

engine = create_engine(
            "mysql+pymysql://{0}:{1}@{2}/{3}?charset=utf8".format(MysqlConfig.user, MysqlConfig.passwd, MysqlConfig.host, MysqlConfig.databaseName))
tdate = pd.read_sql('SELECT * FROM tradedate', engine)

class DealTradeDate():
    # tdate = pd.read_excel(r'E:\myprog\tradedate.xlsx')
    def __init__(self):
        pass

    # TODO 存在的问题是如果输入日期超出了tdate中的日期，则可能返回的是tdate中的最后一个日期，必须及时更新tdate中的日期
    @staticmethod
    def timedate(datet):
        return datetime(datet.year, datet.month, datet.day)

    # tdate=pd.read_excel('./tradedate.xlsx')
    # tdate = np.array([DealTradeDate.timedate(pd.Timestamp(i[0])) for i in tdate.values])

    @staticmethod
    def TradeDateOffset(Date, OffSet):
        """

        :param Date: Date should be '20190612' '2019-6-12' or '2019/6/12'or datetime type,
                     Date don't need to be a trade date
        :param OffSet: integer + means after Date,- means before Date
        :return: Date, Timestamp type
        """
        DateIndex = tdate[tdate.Tdate<=pd.Timestamp(Date)].index[-1]
        NeedDateIndex = DateIndex + OffSet
        NeedDate = tdate.Tdate[NeedDateIndex]
        return NeedDate

    @staticmethod
    def TradateDateCount(StartDate, EndDate):
        """
        :param StartDate: should be '20190612' '2019-6-12' or '2019/6/12'or datetime type
                            Start Date don't need to be a trade date
        :param EndDate: should be '20190612' '2019-6-12' or '2019/6/12'or datetime type
                            End Date don't need to be a trade date
        :return: total days between StartDate and  EndDate, integer
        """
        StartDateIndex = tdate[tdate.Tdate<=pd.Timestamp(StartDate)].index[-1]
        EndDateIndex = tdate[tdate.Tdate <=pd.Timestamp(EndDate)].index[-1]
        CountDays = EndDateIndex - StartDateIndex
        return CountDays

    @staticmethod
    def GetTradeDate(StartDate, EndDate):
        """
        :param StartDate: should be '20190612' '2019-6-12' or '2019/6/12'or datetime type
        :param EndDate: should be '20190612' '2019-6-12' or '2019/6/12'or datetime type
        :return: Trade date series between StartDate and EndDate
        """
        return tdate[(tdate.Tdate>=pd.Timestamp(StartDate))&(tdate.Tdate<=pd.Timestamp(EndDate))]



# print(tdate)