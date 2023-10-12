import pandas as pd
import numpy as np
from XBY.Settle.SettleFuctionGlobal import Settle_Global
from XBY.Settle.ReadCashStatementGlobal import ReadCashStatement
from XBY.Settle.AutoTransfer import AutoTrans
from WindPy import w
w.start()
import os
import itertools


# Dailyreport = pd.read_excel('场外期权业务日报(2020-07-24).xlsx',sheet_name='场外期权合约信息')

def GenNumList(num):
    i = 0
    res =[]
    while i<num:
        res.append(i)
        i += 1
    return res

class CashCheck:
    def __init__(self):
        self.Today_Settle = Settle_Global()
        self.Today_Settle.CheckKnock()
        self.Today_Settle.GetSettleContract()
        self.Today_Settle.CashFLOW()
        self.Today_Settle.LastDayCashFLOW()
        self.Today_Cash_ACTU = ReadCashStatement("{0}{1}".format("{}".format(self.Today_Settle.TDate.month).zfill(2),
                                                       "{}".format(self.Today_Settle.TDate.day).zfill(2)))
        # if
        self.Today_Settle_DF = self.Today_Settle.Settle_DF
        self.Today_Cash = self.Today_Settle.Cash_DF
        self.LastDay_Cash = self.Today_Settle.Last_Cash_DF

        self.Actual_Cash = self.Today_Cash_ACTU.HistoryDetails
        self.Today_file_path = 'Z:\TEST\\'
        self.Today_str_Accrued = 'CashFlow_Accrued_{0}{1}{2}.xlsx'.format("{}".format(self.Today_Settle.TDate.year).zfill(4),
                                                                     "{}".format(self.Today_Settle.TDate.month).zfill(2),
                                                                     "{}".format(self.Today_Settle.TDate.day).zfill(2))
        self.Today_str_Actual = 'CashFlow_Actual_{0}{1}{2}.xlsx'.format("{}".format(self.Today_Settle.TDate.year).zfill(4),
                                                                   "{}".format(self.Today_Settle.TDate.month).zfill(2),
                                                                   "{}".format(self.Today_Settle.TDate.day).zfill(2))

        self.Today_str_Accrued = self.Today_file_path + self.Today_str_Accrued
        self.Today_str_Actual = self.Today_file_path + self.Today_str_Actual

        No_Corr_Accrued_str = 'Nocorr_Accrued.xlsx'
        self.No_Corr_Accrued_str = self.Today_file_path + No_Corr_Accrued_str

        No_Corr_Actual_str = 'Nocorr_Actual.xlsx'
        self.No_Corr_Actual_str = self.Today_file_path + No_Corr_Actual_str

        if os.path.exists(self.No_Corr_Actual_str):
            temp_Actual_Cash = pd.read_excel(self.No_Corr_Actual_str )
            temp_Actual_Cash = temp_Actual_Cash.drop(columns=['Unnamed: 0'])
            temp_Accrued_Cash = pd.read_excel(self.No_Corr_Accrued_str)
            temp_Accrued_Cash = temp_Accrued_Cash.drop(columns=['Unnamed: 0'])

            self.Actual_Cash = self.Actual_Cash.append(temp_Actual_Cash)
            self.Today_Cash = self.Today_Cash.append(temp_Accrued_Cash)

        self.Today_Cash = self.Today_Cash.append(self.LastDay_Cash)


        self.Today_Cash.to_excel(self.Today_str_Accrued)
        self.Actual_Cash.to_excel(self.Today_str_Actual)

    def Add_before(self):


        self.Accrued = pd.read_excel(self.Today_str_Accrued)
        self.Actual = pd.read_excel(self.Today_str_Actual)
        self.Accrued = self.Accrued.drop(columns=['Unnamed: 0'])
        self.Actual = self.Actual.drop(columns=['Unnamed: 0'])
        self.Accrued['对应'] = 0
        self.Accrued['时间'] = 0
        self.Actual['对应'] = 0
        self.Actual['对应交易'] = 0
        self.Actual['差额'] = 0


        temp_index = self.Accrued[self.Accrued['汇总'] == 0].index.values
        print('AAAAAA',temp_index)
        self.Accrued = self.Accrued.drop(temp_index)



    def CheckCorrs(self):

        Accrued_CounterParty = list(self.Accrued['交易对手'].drop_duplicates().values)
        Actual_CounterParty = list(self.Actual['交易对手'].drop_duplicates().values)
        Bothin = [x for x in Accrued_CounterParty if x in Actual_CounterParty]
        Diff = [y for y in (Accrued_CounterParty + Actual_CounterParty) if y not in Bothin]


        for CounterParty in Bothin:
            Slice_Accrued = self.Accrued[self.Accrued['交易对手'] == CounterParty]
            Slice_Actual = self.Actual[self.Actual['交易对手'] == CounterParty]
            print(Slice_Accrued['交易对手'])

            if CounterParty in ['江南农商行','招商银行']:
                # temp_len_Actual = len(Slice_Actual)

                for cashflow_Actual in list(Slice_Actual['发生额']):
                    if cashflow_Actual > 0:
                        if self.Accrued[self.Accrued['期权费'] == cashflow_Actual].empty:
                            pass
                        else:
                            index_Accrued = self.Accrued[self.Accrued['期权费'] == cashflow_Actual].index[0]
                            index_Actual = self.Actual[self.Actual['发生额'] == cashflow_Actual].index[0]
                            self.Accrued['对应'][index_Accrued] = 1

                            self.Accrued['时间'][index_Accrued] = self.Actual['交易时间'][index_Actual]
                            self.Actual['对应交易'][index_Actual] = self.Accrued['交易编号'].values[index_Accrued]
                            self.Actual['对应'][index_Actual] = 1
                    else:
                        temp_len_Accrued = len(Slice_Accrued)

                        i = 0
                        while i < temp_len_Accrued:
                            temp_payoff = []
                            temp_check = list(Slice_Accrued['行权收益'].values)[i]
                            temp_ContractId = list(Slice_Accrued['交易编号'].values)
                            temp_payoff.append(temp_check)

                            if temp_check == cashflow_Actual:
                                index_Actual = self.Actual[self.Actual['发生额'] == cashflow_Actual].index[0]
                                index_Accrued = self.Accrued[self.Accrued['行权收益'] == cashflow_Actual].index[0]
                                self.Actual['对应'][index_Actual] = 1
                                str_temp_ContractId = ','.join(temp_ContractId)

                                self.Accrued['时间'][index_Accrued] = self.Actual['交易时间'][index_Actual]
                                self.Actual['对应交易'][index_Actual] = str_temp_ContractId
                                # if self.Accrued['时间'][index_Accrued] == 0:
                                #     raise ('未付行权收益')

                            else:
                                for j in range(i + 1, temp_len_Accrued):
                                    temp_check = round(temp_check + list(Slice_Accrued['行权收益'].values)[j], 2)
                                    temp_payoff.append(list(Slice_Accrued['行权收益'].values)[j])
                                    if temp_check == cashflow_Actual:
                                        index_Actual = self.Actual[self.Actual['发生额'] == cashflow_Actual].index[0]
                                        self.Actual['对应'][index_Actual] = 1
                                        str_temp_ContractId = ','.join(temp_ContractId)

                                        self.Actual['对应交易'][index_Actual] = str_temp_ContractId
                                        len_temp_payoff = len(temp_payoff)
                                        for len_temp_index_Slice in range(len_temp_payoff):
                                            index_Accrued = self.Accrued[self.Accrued['行权收益'] == temp_payoff[len_temp_index_Slice]].index[0]
                                            # if self.Accrued['时间'][index_Accrued] == 0:
                                            #     raise ('未付行权收益')

                            i = i + 1

            temp_len_Accrued = len(Slice_Accrued)
            temp_list = GenNumList(temp_len_Accrued)
            for cashflow_Actual in list(Slice_Actual['发生额']):
                # print(cashflow_Actual)
                print(temp_list)
                i = 1
                if CounterParty == '中金公司':
                    len_loop = 4
                elif CounterParty == '中信证券':
                    len_loop = 4
                elif CounterParty == '广金美好基金':
                    len_loop = 2
                else:
                    len_loop = 100
                while i < min(temp_len_Accrued + 1, len_loop):

                    CheckList = list(itertools.combinations(temp_list, i))
                    i += 1
                    for k in range(len(CheckList)):

                        temp_index_pop = list(CheckList[k])

                        temp_payoff = list(Slice_Accrued['汇总'].values[list(CheckList[k])])
                        temp_ContractId = list(Slice_Accrued['交易编号'].values[list(CheckList[k])])
                        len_temp_payoff = len(temp_payoff)
                        temp_check = round(sum(temp_payoff), 2)

                        print(CheckList[k])

                        if temp_check == cashflow_Actual:

                            temp_Actual = self.Actual[self.Actual['交易对手'] == CounterParty]
                            index_Actual = temp_Actual[temp_Actual['发生额'] == cashflow_Actual].index[0]
                            self.Actual['对应'][index_Actual] = 1
                            str_temp_ContractId = ','.join(temp_ContractId)
                            self.Actual['对应交易'][index_Actual] = str_temp_ContractId

                            for len_temp_index_Slice in range(len_temp_payoff):
                                temp_Accrued = self.Accrued[self.Accrued['交易对手'] == CounterParty]
                                index_Accrued = temp_Accrued[
                                    temp_Accrued['汇总'] == temp_payoff[len_temp_index_Slice]].index.values
                                self.Accrued['对应'][index_Accrued] = 1
                                self.Accrued['时间'][index_Accrued] = self.Actual['交易时间'][index_Actual]


                            for index in temp_index_pop:
                                temp_list.remove(index)
                            print(str_temp_ContractId, '精确匹配')
                            break

                        NO_ACCURATE = list(np.linspace(cashflow_Actual - 50, cashflow_Actual + 50, 10001))
                        NO_ACCURATE = list(map(lambda x: round(x, 2), NO_ACCURATE))
                        NO_ACCURATE.remove(cashflow_Actual)
                        if temp_check in NO_ACCURATE:
                            temp_Actual = self.Actual[self.Actual['交易对手'] == CounterParty]
                            index_Actual = temp_Actual[temp_Actual['发生额'] == cashflow_Actual].index[0]
                            self.Actual['对应'][index_Actual] = 2
                            str_temp_ContractId = ','.join(temp_ContractId)
                            self.Actual['对应交易'][index_Actual] = str_temp_ContractId
                            for len_temp_index_Slice in range(len_temp_payoff):
                                temp_Accrued = self.Accrued[self.Accrued['交易对手'] == CounterParty]
                                index_Accrued = temp_Accrued[
                                    temp_Accrued['汇总'] == temp_payoff[len_temp_index_Slice]].index.values
                                self.Accrued['对应'][index_Accrued] = 2
                                self.Accrued['时间'][index_Accrued] = self.Actual['交易时间'][index_Actual]

                                print(str_temp_ContractId, '差额', round(temp_check - cashflow_Actual, 2))
                            self.Actual['差额'][index_Actual] = round(temp_check - cashflow_Actual, 2)
                            for index in temp_index_pop:
                                temp_list.remove(index)
                            print(str_temp_ContractId, '精确匹配')
                            break


        for CounterParty in Bothin:
            Slice_Accrued = self.Accrued[self.Accrued['交易对手'] == CounterParty]
            Slice_Accrued = Slice_Accrued[Slice_Accrued['对应'] == 0]
            Slice_Actual = self.Actual[self.Actual['交易对手'] == CounterParty]
            Slice_Actual = Slice_Actual[Slice_Actual['对应'] == 0]

            if Slice_Accrued.empty | Slice_Actual.empty == False:

                print(Slice_Accrued['交易对手'])

                for cashflow_Accruted in list(Slice_Accrued['汇总']):
                    temp_len_Actural = len(Slice_Actual)
                    temp_list = GenNumList(temp_len_Actural)
                    # print(cashflow_Actual)
                    print(temp_list)

                    i = 1
                    if CounterParty == '中金公司':
                        len_loop = 3
                    else:
                        len_loop = 3
                    while i < min(temp_len_Actural + 1, len_loop):

                        CheckList = list(itertools.combinations(temp_list, i))
                        i += 1
                        for k in range(len(CheckList)):

                            temp_index_pop = list(CheckList[k])

                            temp_payoff = list(Slice_Actual['发生额'].values[list(CheckList[k])])
                            temp_time = list(Slice_Actual['交易时间'].values[list(CheckList[k])])
                            len_temp_payoff = len(temp_payoff)
                            temp_check = round(sum(temp_payoff), 2)

                            print(CheckList[k])

                            if temp_check == cashflow_Accruted:

                                temp_Accrued = self.Accrued[self.Accrued['交易对手'] == CounterParty]
                                index_Accrued = temp_Accrued[temp_Accrued['汇总'] == cashflow_Accruted].index[0]
                                self.Accrued['对应'][index_Accrued] = 1
                                str_temp_time = ','.join(temp_time)
                                self.Accrued['时间'][index_Accrued] = str_temp_time

                                for len_temp_index_Slice in range(len_temp_payoff):
                                    index_Actual = self.Actual[
                                        self.Actual['发生额'] == temp_payoff[len_temp_index_Slice]].index.values
                                    if self.Actual['交易对手'][index_Actual].values[0] == CounterParty:
                                        self.Actual['对应'][index_Actual] = 1
                                        self.Actual['对应交易'][index_Actual] = self.Accrued['交易编号'][index_Accrued]

                                for index in temp_index_pop:
                                    temp_list.remove(index)
                                print(self.Accrued['交易编号'][index_Accrued], '精确匹配')
                                break

                            NO_ACCURATE = list(np.linspace(cashflow_Accruted - 50, cashflow_Accruted + 50, 10001))
                            NO_ACCURATE = list(map(lambda x: round(x, 2), NO_ACCURATE))
                            print(cashflow_Accruted)
                            NO_ACCURATE.remove(round(cashflow_Accruted,2))
                            if temp_check in NO_ACCURATE:

                                # index_Accrued = self.Accrued[self.Accrued['汇总'] == cashflow_Accruted].index[0]

                                temp_Accrued = self.Accrued[self.Accrued['交易对手'] == CounterParty]
                                index_Accrued = temp_Accrued[temp_Accrued['汇总'] == cashflow_Accruted].index[0]
                                self.Accrued['对应'][index_Accrued] = 1
                                str_temp_time = ','.join(temp_time)
                                self.Accrued['时间'][index_Accrued] = str_temp_time

                                for len_temp_index_Slice in range(len_temp_payoff):
                                    index_Actual = self.Actual[
                                        self.Actual['发生额'] == temp_payoff[len_temp_index_Slice]].index.values
                                    if self.Actual['交易对手'][index_Actual].values[0] == CounterParty:
                                        self.Actual['对应'][index_Actual] = 2
                                        self.Actual['对应交易'][index_Actual] = self.Accrued['交易编号'][index_Accrued]

                                    print(self.Accrued['交易编号'][index_Accrued], '差额',
                                          round(temp_check - cashflow_Accruted, 2))
                                self.Actual['差额'][index_Actual] = round(temp_check - cashflow_Accruted, 2)
                                for index in temp_index_pop:
                                    temp_list.remove(index)
                                print(self.Accrued['交易编号'][index_Accrued], '精确匹配')
                                break

        # for CounterParty in Bothin:
        #     Slice_Accrued = self.Accrued[self.Accrued['交易对手'] == CounterParty]
        #     Slice_Accrued = Slice_Accrued[Slice_Accrued['对应'] == 0]
        #     Slice_Actual = self.Actual[self.Actual['交易对手'] == CounterParty]
        #     Slice_Actual = Slice_Actual[Slice_Actual['对应'] == 0]
        #
        #     if Slice_Accrued.empty | Slice_Actual.empty == False:


    def NoCorrToExcel(self):
        self.No_Corr_Accrued = self.Accrued[self.Accrued['对应'] == 0]
        self.No_Corr_Accrued = self.No_Corr_Accrued.drop(['对应', '时间'], 1)
        self.No_Corr_Accrued.to_excel(self.No_Corr_Accrued_str)

        self.No_Corr_Actual = self.Actual[self.Actual['对应'] == 0]
        self.No_Corr_Actual = self.No_Corr_Actual.drop(['对应'], 1)
        self.No_Corr_Actual.to_excel(self.No_Corr_Actual_str)

        self.Accrued.to_excel(self.Today_str_Accrued)
        self.Actual.to_excel(self.Today_str_Actual)

    def AutoTransfer(self):
        AutoTrans(self.Today_Settle.TDate.strftime('%Y%m%d'))


A = CashCheck()
A.Add_before()
A.CheckCorrs()
A.NoCorrToExcel()
A.AutoTransfer()


