from datetime import datetime
import pandas as pd
import openpyxl
from sqlalchemy import create_engine

from WindPy import w
from TradeDate import DealTradeDate
# from KDBGenerate_Global import KDB_Global
from MyReport import ExoticReport


class DailyWorK:
    def __init__(self, TDate = ''):
        self.FilePath = 'E.\Contracts_and_trade\MyTotalContracts.xlsx'
        self.mainopinfo = pd.read_excel(self.FilePath)
        today = datetime.today()
        if TDate=='':
            if int(today.strftime('%H%M'))>1530:
                TDate = DealTradeDate.timedate(w.tdaysoffset(0, today).Data[0][0])
            else:
                if DealTradeDate.timedate(w.tdaysoffset(0, today).Data[0][0]).strftime('%Y%m%d')==today.strftime('%Y%m%d'):
                    TDate = DealTradeDate.timedate(w.tdaysoffset(-1,today).Data[0][0])
                else:
                    TDate = DealTradeDate.timedate(w.tdaysoffset(0, today).Data[0][0])
        self.TDate = TDate
        self.TDate_next = DealTradeDate.timedate(w.tdaysoffset(1,TDate).Data[0][0])

    def DecomposeReport(self):
        print("start make OTC_PL_Decompose_Summary")

        mainopinfo = self.mainopinfo[self.mainopinfo.ActualEndDate>=self.TDate]
        UnderlyingSet = list(set(mainopinfo.Underlying))
        UnderlyingSet = [i for i in UnderlyingSet if (i!='CorrMax')&(i!='CorrPro3')&(i!='002001.SZ')]
        UnderlyingSet = ['ZZ500']
        for uncode in UnderlyingSet:
            Re = ExoticReport(uncode,self.TDate)
            # Re.GetKnockIn()
            # # Re.GetTouchExpire()
            Re.genreport()

        # Re = ExoticReport('AU', self.TDate)
        # Re.genreport()
        print("OTC_PL_Decompose_Summary made done")

