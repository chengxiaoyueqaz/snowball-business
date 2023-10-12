# -*- coding: utf-8 -*-
# @Author  : chenyihui
# @Email   : chenyihui10@126.com
# @Time    :2018-9-3 18:19
"""新结构需要另外optionfunction添加一个函数，新起息时会将全部希腊字母计算出来再使用其中的价格"""
import numpy as np

# import optionMatlabfunction as of
# import optionfunction as of
import Pricing.optionfunction as of
from KDBGenerate_Global import Get_Discounted_Ratio, GetStrSplit
from SnowballPDE.AutoCall_Prob_SnowBall_2 import AutoCall_SnowBall_MyReport
from SnowballPDE.AutoCall_Step_SnowBall_Prob import AutoCall_SnowBall_Step_MyReport
from SnowballPDE.NewSnowBall import AutoCall_NewSnowBall_MyReport
YearDays = 250

DepositRate = 1.5

HedgeCode = {'HS300': 'IF2106.CFE',
             'ZZ500': 'IC2106.CFE',
             '1833.HK': '1833.HK',
             '0772.HK': '0772.HK',}

class ReportInfo:
    YearDays = 250

    DepositRate = 1.5

    # OptionDict = {
    #     u'欧式看涨': lambda x1, x2, x3, x4, x5, x6, x7: of.VanillaCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                                                      x6),
    #     u'美式看涨': lambda x1, x2, x3, x4, x5, x6, x7: of.VanillaCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                                                      x6),
    #     u'备兑看涨': lambda x1, x2, x3, x4, x5, x6, x7: of.VanillaCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                                                      x6),
    #     u'欧式看跌': lambda x1, x2, x3, x4, x5, x6, x7: of.VanillaPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5,
    #                                                                     x6),
    #     u'美式看跌': lambda x1, x2, x3, x4, x5, x6, x7: of.VanillaPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5,
    #                                                                     x6),
    #     u'看涨价差': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x6) -
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x6)),
    #     u'看跌价差': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x6) -
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x6)),
    #     u'欧式二元看涨': lambda x1, x2, x3, x4, x5, x6, x7: of.BinaryCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                                                       x7.InitialS * x7.Rebate, x6),
    #     u'欧式二元看跌': lambda x1, x2, x3, x4, x5, x6, x7: of.BinaryPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5,
    #                                                                      x7.InitialS * x7.Rebate, x6),
    #     u'双向鲨鱼鳍': lambda x1, x2, x3, x4, x5, x6, x7: of.DoubleSharkfinGreeks(x1, x7.InitialS * x7.LstrikeRatio,
    #                                                                          x7.InitialS * x7.HstrikeRatio,
    #                                                                          x2, x3, x4, x5, x7.InitialS * x7.LtouchRatio,
    #                                                                          x7.InitialS * x7.HtouchRatio,
    #                                                                          x7.InitialS * x7.Rebate, x6),
    #     u'不对称双向鲨鱼鳍': lambda x1, x2, x3, x4, x5, x6, x7: of.NoEqualDoubleSharkfinGreeks(x1, x7.InitialS * x7.LstrikeRatio,
    #                                                                                    x7.InitialS * x7.HstrikeRatio,
    #                                                                                    x2, x3, x4, x5,
    #                                                                                    x7.InitialS * x7.LtouchRatio,
    #                                                                                    x7.InitialS * x7.HtouchRatio,
    #                                                                                    x7.InitialS * x7.Rebate,
    #                                                                                    x7.InitialS * x7.Rebate2, x6),
    #     u'单向鲨鱼鳍看涨': lambda x1, x2, x3, x4, x5, x6, x7: of.UpOutCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                                                       x7.InitialS * x7.HtouchRatio,
    #                                                                       x7.InitialS * x7.Rebate, x6),
    #     u'单向鲨鱼鳍看跌': lambda x1, x2, x3, x4, x5, x6, x7: of.DownOutPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4,
    #                                                                        x5, x7.InitialS * x7.LtouchRatio,
    #                                                                        x7.InitialS * x7.Rebate, x6),
    #     u'四层累积看涨': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate,
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6) +
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                      x7.InitialS * (x7.Rebate2 - x7.Rebate),
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6) +
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5,
    #                                      x7.InitialS * (x7.Rebate3 - x7.Rebate2),
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6)),
    #     u'三层累积看涨': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate,
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6) +
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                      x7.InitialS * (x7.Rebate2 - x7.Rebate),
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6)),
    #     u'区间累积': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate,
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6) -
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate,
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6)),
    #     u'累积看跌': lambda x1, x2, x3, x4, x5, x6, x7: of.AccrualPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5,
    #                                                                     x7.InitialS * x7.Rebate,
    #                                                                     int(of.mytdate(x7.BeginDate,
    #                                                                                    x7.EndDate) * YearDays) + 1, x6),
    #     u'美式双触碰': lambda x1, x2, x3, x4, x5, x6, x7: of.DoubleSharkfinGreeks(x1, x7.InitialS * x7.LtouchRatio,
    #                                                                          x7.InitialS * x7.HtouchRatio,
    #                                                                          x2, x3, x4, x5, x7.InitialS * x7.LtouchRatio,
    #                                                                          x7.InitialS * x7.HtouchRatio,
    #                                                                          x7.InitialS * x7.Rebate, x6),
    #     u'美式双不触碰': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.AmDoubleNoTouchGreeks(x1, x2, x3, x4, x5, x7.InitialS * x7.LtouchRatio, x7.InitialS * x7.HtouchRatio,
    #                                  x7.InitialS * x7.Rebate, x6)),
    #     u'美式二元看涨': lambda x1, x2, x3, x4, x5, x6, x7: of.UpOutCallGreeks(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5,
    #                                                                      x7.InitialS * x7.HtouchRatio,
    #                                                                      x7.InitialS * x7.Rebate, x6),
    #     u'美式二元看涨最高价': lambda x1, x2, x3, x4, x5, x6, x7: of.UpOutCallGreeks(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4,
    #                                                                         x5, x7.InitialS * x7.HtouchRatio,
    #                                                                         x7.InitialS * x7.Rebate, x6),
    #     u'美式二元看跌': lambda x1, x2, x3, x4, x5, x6, x7: of.DownOutPutGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5,
    #                                                                       x7.InitialS * x7.LtouchRatio,
    #                                                                       x7.InitialS * x7.Rebate, x6),
    #     u'美式二元看跌最低价': lambda x1, x2, x3, x4, x5, x6, x7: of.DownOutPutGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4,
    #                                                                          x5, x7.InitialS * x7.LtouchRatio,
    #                                                                          x7.InitialS * x7.Rebate, x6),
    #     u'欧式双向鲨鱼鳍': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x6) -
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5, x6) +
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x6) -
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5, x6) -
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5,
    #                                     x7.InitialS * (x7.HtouchRatio - x7.HstrikeRatio - x7.Rebate), x6) -
    #                 of.BinaryPutGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5,
    #                                    x7.InitialS * (x7.LstrikeRatio - x7.LtouchRatio - x7.Rebate), x6)),
    #     u'二元凸式': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate, x6) -
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate, x6)),
    #     u'四层阶梯看涨': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate, x6) +
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                     x7.InitialS * (x7.Rebate2 - x7.Rebate), x6) +
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5,
    #                                     x7.InitialS * (x7.Rebate3 - x7.Rebate2), x6)),
    #     u'三层阶梯看涨': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate, x6) +
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                     x7.InitialS * (x7.Rebate2 - x7.Rebate), x6)),
    #     u'三层阶梯看跌': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.BinaryPutGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate, x6) +
    #                 of.BinaryPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5,
    #                                    x7.InitialS * (x7.Rebate2 - x7.Rebate), x6)),
    #     u'四层阶梯看跌': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.BinaryPutGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate, x6) +
    #                 of.BinaryPutGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                    x7.InitialS * (x7.Rebate2 - x7.Rebate), x6) +
    #                 of.BinaryPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5,
    #                                    x7.InitialS * (x7.Rebate3 - x7.Rebate2), x6)),
    #     u'欧式单向鲨鱼鳍看涨': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x6) -
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5, x6) -
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5,
    #                                     x7.InitialS * (x7.HtouchRatio - x7.HstrikeRatio - x7.Rebate), x6)),
    #     u'欧式单向鲨鱼鳍看跌': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x6) -
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5, x6) -
    #                 of.BinaryPutGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5,
    #                                    x7.InitialS * (x7.LstrikeRatio - x7.LtouchRatio - x7.Rebate), x6)),
    #     u'合成多头': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x6) -
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x6)),
    #     u'二元凹式': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.BinaryPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate, x6) +
    #                 of.BinaryCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate, x6)),
    #     u'四层区间累积': lambda x1, x2, x3, x4, x5, x6, x7: (
    #             of.AccrualCallGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate,
    #                                  int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6) +
    #             of.AccrualCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                  x7.InitialS * (x7.Rebate2 - x7.Rebate),
    #                                  int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6) +
    #             of.AccrualCallGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5,
    #                                  x7.InitialS * (x7.Rebate3 - x7.Rebate2),
    #                                  int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6)),
    #     u'五层区间累积': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate,
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6) +
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                      x7.InitialS * (x7.Rebate2 - x7.Rebate),
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6) +
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5,
    #                                      x7.InitialS * (x7.Rebate3 - x7.Rebate2),
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6) +
    #                 of.AccrualCallGreeks(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5,
    #                                      x7.InitialS * (x7.Rebate4 - x7.Rebate3),
    #                                      int(of.mytdate(x7.BeginDate, x7.EndDate) * YearDays) + 1, x6)),
    #     u'跨式组合': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x6) * x7.Pratio1 +
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x6) * x7.Pratio2),
    #     u'鹰式组合': lambda x1, x2, x3, x4, x5, x6, x7: (
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5, x6) -
    #                 of.VanillaCallGreeks(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5, x6) +
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.LstrikeRatio, x2, x3, x4, x5, x6) -
    #                 of.VanillaPutGreeks(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5, x6)),
    #     u'看涨自动敲出赎回': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.AutoCallEntry(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate,
    #                          x7.InitialS * x7.Premium / x7.Pratio, x7, x6)),
    #     u'看涨自动敲出赎回新起息': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.AutoCallNewEntry(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5,
    #                             x7.InitialS * x7.Rebate, x7.InitialS * x7.Premium / x7.Pratio, x7, x6)),
    #     u'敲入看涨自动敲出赎回': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.KnockInAutoCallEntry(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5, x7.InitialS * x7.HstrikeRatio,
    #                                 x7.InitialS * x7.Rebate, x7.InitialS * x7.Premium / x7.Pratio, x7, x6)),
    #     u'敲入看涨自动敲出赎回新起息': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.KnockInAutoCallNewEntry(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5, x7.InitialS * x7.HstrikeRatio,
    #                                    x7.InitialS * x7.Rebate, x7.InitialS * x7.Premium / x7.Pratio, x7, x6)),
    #     u'逐步调整看涨自动敲出赎回': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.AutoCallStepChangeEntry(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate3,
    #                                    x7.InitialS * x7.Rebate, x7.InitialS * x7.Premium / x7.Pratio, x7, x6)),
    #     u'逐步调整看涨自动敲出赎回新起息': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.AutoCallStepChangeNewEntry(x1, x7.InitialS * x7.HtouchRatio, x2, x3, x4, x5,
    #                                       x7.InitialS * x7.Rebate3, x7.InitialS * x7.Rebate,
    #                                       x7.InitialS * x7.Premium / x7.Pratio, x7, x6)),
    #     u'看跌自动敲出赎回': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.AutoPutEntry(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5, x7.InitialS * x7.Rebate,
    #                         x7.InitialS * x7.Premium / x7.Pratio, x7, x6)),
    #     u'看跌自动敲出赎回新起息': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.AutoPutNewEntry(x1, x7.InitialS * x7.LtouchRatio, x2, x3, x4, x5,
    #                            x7.InitialS * x7.Rebate, x7.InitialS * x7.Premium / x7.Pratio, x7, x6)),
    #     u'看涨自动敲出赎回转看涨': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.AutoCallableToCallEntry(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                    x7.InitialS * x7.HtouchRatio, x7.InitialS * x7.Rebate,
    #                                    x7.InitialS * x7.Premium / x7.Pratio, x7, x6)),
    #     u'看涨自动敲出赎回转看涨新起息': lambda x1, x2, x3, x4, x5, x6, x7: (
    #         of.AutoCallableToCallNewEntry(x1, x7.InitialS * x7.HstrikeRatio, x2, x3, x4, x5,
    #                                       x7.InitialS * x7.HtouchRatio, x7.InitialS * x7.Rebate,
    #                                       x7.InitialS * x7.Premium / x7.Pratio, x7, x6))}
    OptionDict = {
        u'欧式看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.VanillaCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf, Dividend,
            GreeksType),
        u'指数增强型收益互换': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.SwapGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, GreeksType)*(
            Get_Discounted_Ratio(ContractInfo.Underlying, HedgeCode[ContractInfo.Underlying]) if 'Price' not in GreeksType else 1)*(
            np.array([1,Get_Discounted_Ratio(ContractInfo.Underlying, HedgeCode[ContractInfo.Underlying]),1,1,1]) if GreeksType=='All' else 1
        ),
        u'美式看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.VanillaCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf, Dividend,
            GreeksType),
        u'备兑看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.VanillaCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf, Dividend,
            GreeksType),
        u'欧式看跌': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.VanillaPutGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf, Dividend,
            GreeksType),
        u'美式看跌': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.VanillaPutGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf, Dividend,
            GreeksType),
        u'看涨价差': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) -
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType)),
        u'区间保护': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) -
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) -
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) +
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType)),
        u'看跌价差': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType) -
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType)),
        u'欧式二元看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.BinaryCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf, Dividend,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'欧式二元看跌': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.BinaryPutGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf, Dividend,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'双向鲨鱼鳍': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.DoubleSharkfinGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio,
            ContractInfo.InitialS * ContractInfo.HstrikeRatio,
            Maturity, Vol, Rf, Dividend, ContractInfo.InitialS * ContractInfo.LtouchRatio,
            ContractInfo.InitialS * ContractInfo.HtouchRatio,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'不对称双向鲨鱼鳍': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType,
                            ContractInfo: of.NoEqualDoubleSharkfinGreeks(PricingDate,
                                                                         ContractInfo.InitialS * ContractInfo.LstrikeRatio,
                                                                         ContractInfo.InitialS * ContractInfo.HstrikeRatio,
                                                                         Maturity, Vol, Rf, Dividend,
                                                                         ContractInfo.InitialS * ContractInfo.LtouchRatio,
                                                                         ContractInfo.InitialS * ContractInfo.HtouchRatio,
                                                                         ContractInfo.InitialS * ContractInfo.Rebate,
                                                                         ContractInfo.InitialS * ContractInfo.Rebate2,
                                                                         GreeksType),
        u'单向鲨鱼鳍看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.UpOutCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf, Dividend,
            ContractInfo.InitialS * ContractInfo.HtouchRatio,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'自动赎回看涨价差': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (of.UpOutCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf, Dividend,
                         ContractInfo.InitialS * ContractInfo.HtouchRatio,
                         ContractInfo.InitialS * (ContractInfo.HtouchRatio-ContractInfo.LstrikeRatio), GreeksType)*ContractInfo.Pratio1 -
            of.UpOutCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf, Dividend,
            ContractInfo.InitialS * ContractInfo.HtouchRatio,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType)*(ContractInfo.Pratio1-ContractInfo.Pratio2)),
        u'单向鲨鱼鳍看跌': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.DownOutPutGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
            Dividend, ContractInfo.InitialS * ContractInfo.LtouchRatio,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'四层累积看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, ContractInfo.InitialS * ContractInfo.Rebate,
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType) +
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend,
                                     ContractInfo.InitialS * (ContractInfo.Rebate2 - ContractInfo.Rebate),
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType) +
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                     Dividend,
                                     ContractInfo.InitialS * (ContractInfo.Rebate3 - ContractInfo.Rebate2),
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType)),
        u'三层累积看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, ContractInfo.InitialS * ContractInfo.Rebate,
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType) +
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend,
                                     ContractInfo.InitialS * (ContractInfo.Rebate2 - ContractInfo.Rebate),
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType)),
        u'区间累积': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, ContractInfo.InitialS * ContractInfo.Rebate,
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType) -
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, ContractInfo.InitialS * ContractInfo.Rebate,
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType)),
        u'累积看跌': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.AccrualPutGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf, Dividend,
            ContractInfo.InitialS * ContractInfo.Rebate,
            int(of.mytdate(ContractInfo.BeginDate,
                           ContractInfo.EndDate) * YearDays) + 1, GreeksType),
        u'美式双触碰': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.DoubleSharkfinGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio,
            ContractInfo.InitialS * ContractInfo.HtouchRatio,
            Maturity, Vol, Rf, Dividend, ContractInfo.InitialS * ContractInfo.LtouchRatio,
            ContractInfo.InitialS * ContractInfo.HtouchRatio,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'美式双不触碰': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.AmDoubleNoTouchGreeks(PricingDate, Maturity, Vol, Rf, Dividend,
                                     ContractInfo.InitialS * ContractInfo.LtouchRatio,
                                     ContractInfo.InitialS * ContractInfo.HtouchRatio,
                                     ContractInfo.InitialS * ContractInfo.Rebate, GreeksType)),
        u'美式二元看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.UpOutCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf, Dividend,
            ContractInfo.InitialS * ContractInfo.HtouchRatio,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'美式二元向上不触碰': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo:
        of.AmSingleNoTouchCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf, Dividend,
                         ContractInfo.InitialS * ContractInfo.HtouchRatio,
                         ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'美式二元看涨最高价': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.UpOutCallGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
            Dividend, ContractInfo.InitialS * ContractInfo.HtouchRatio,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'美式二元看跌': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.DownOutPutGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf, Dividend,
            ContractInfo.InitialS * ContractInfo.LtouchRatio,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'美式二元看跌最低价': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: of.DownOutPutGreeks(
            PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
            Dividend, ContractInfo.InitialS * ContractInfo.LtouchRatio,
            ContractInfo.InitialS * ContractInfo.Rebate, GreeksType),
        u'欧式双向鲨鱼鳍': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) -
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) +
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType) -
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType) -
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                    Dividend,
                                    ContractInfo.InitialS * (
                                                ContractInfo.HtouchRatio - ContractInfo.HstrikeRatio - ContractInfo.Rebate),
                                    GreeksType) -
                of.BinaryPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                   Dividend,
                                   ContractInfo.InitialS * (
                                               ContractInfo.LstrikeRatio - ContractInfo.LtouchRatio - ContractInfo.Rebate),
                                   GreeksType)),
        u'二元凸式': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, ContractInfo.InitialS * ContractInfo.Rebate, GreeksType) -
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, ContractInfo.InitialS * ContractInfo.Rebate, GreeksType)),
        u'四层阶梯看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, ContractInfo.InitialS * ContractInfo.Rebate, GreeksType) +
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                    Dividend,
                                    ContractInfo.InitialS * (ContractInfo.Rebate2 - ContractInfo.Rebate), GreeksType) +
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                    Dividend,
                                    ContractInfo.InitialS * (ContractInfo.Rebate3 - ContractInfo.Rebate2), GreeksType)),
        u'三层阶梯看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, ContractInfo.InitialS * ContractInfo.Rebate, GreeksType) +
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                    Dividend,
                                    ContractInfo.InitialS * (ContractInfo.Rebate2 - ContractInfo.Rebate), GreeksType)),
        u'三层阶梯看跌': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.BinaryPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                   Dividend, ContractInfo.InitialS * ContractInfo.Rebate, GreeksType) +
                of.BinaryPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                   Dividend,
                                   ContractInfo.InitialS * (ContractInfo.Rebate2 - ContractInfo.Rebate), GreeksType)),
        u'四层阶梯看跌': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.BinaryPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                   Dividend, ContractInfo.InitialS * ContractInfo.Rebate, GreeksType) +
                of.BinaryPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                   Dividend,
                                   ContractInfo.InitialS * (ContractInfo.Rebate2 - ContractInfo.Rebate), GreeksType) +
                of.BinaryPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                   Dividend,
                                   ContractInfo.InitialS * (ContractInfo.Rebate3 - ContractInfo.Rebate2), GreeksType)),
        u'欧式单向鲨鱼鳍看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) -
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) -
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                    Dividend,
                                    ContractInfo.InitialS * (
                                                ContractInfo.HtouchRatio - ContractInfo.HstrikeRatio - ContractInfo.Rebate),
                                    GreeksType)),
        u'欧式单向鲨鱼鳍看跌': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType) -
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType) -
                of.BinaryPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                   Dividend,
                                   ContractInfo.InitialS * (
                                               ContractInfo.LstrikeRatio - ContractInfo.LtouchRatio - ContractInfo.Rebate),
                                   GreeksType)),
        u'合成多头': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) -
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType)),
        u'二元凹式': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.BinaryPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                   Dividend, ContractInfo.InitialS * ContractInfo.Rebate, GreeksType) +
                of.BinaryCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, ContractInfo.InitialS * ContractInfo.Rebate, GreeksType)),
        u'四层区间累积': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, ContractInfo.InitialS * ContractInfo.Rebate,
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType) +
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend,
                                     ContractInfo.InitialS * (ContractInfo.Rebate2 - ContractInfo.Rebate),
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType) +
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                     Dividend,
                                     ContractInfo.InitialS * (ContractInfo.Rebate3 - ContractInfo.Rebate2),
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType)),
        u'五层区间累积': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, ContractInfo.InitialS * ContractInfo.Rebate,
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType) +
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend,
                                     ContractInfo.InitialS * (ContractInfo.Rebate2 - ContractInfo.Rebate),
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType) +
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                     Dividend,
                                     ContractInfo.InitialS * (ContractInfo.Rebate3 - ContractInfo.Rebate2),
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType) +
                of.AccrualCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                     Dividend,
                                     ContractInfo.InitialS * (ContractInfo.Rebate4 - ContractInfo.Rebate3),
                                     int(of.mytdate(ContractInfo.BeginDate, ContractInfo.EndDate) * YearDays) + 1,
                                     GreeksType)),
        u'跨式组合': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) * ContractInfo.Pratio1 +
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType) * ContractInfo.Pratio2),
        u'鹰式组合': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) -
                of.VanillaCallGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                     Dividend, GreeksType) +
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LstrikeRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType) -
                of.VanillaPutGreeks(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                                    Dividend, GreeksType)),
        u'看涨自动敲出赎回': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.AutoCallEntry(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf, Dividend,
                             ContractInfo.InitialS * ContractInfo.Rebate,
                             ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio, ContractInfo,
                             GreeksType)),
        u'雪球看涨自动敲出赎回': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            AutoCall_SnowBall_MyReport(PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo)),
        u'逐步调整雪球看涨自动敲出赎回': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            AutoCall_SnowBall_Step_MyReport(PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo)),
        u'收益凭证雪球': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            AutoCall_NewSnowBall_MyReport(PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo)),
        u'看涨自动敲出赎回新起息': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.AutoCallNewEntry(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                Dividend,
                                ContractInfo.InitialS * ContractInfo.Rebate,
                                ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio, ContractInfo,
                                GreeksType)),
        u'敲入看涨自动敲出赎回': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.KnockInAutoCallEntry(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                    Dividend, ContractInfo.InitialS * ContractInfo.HstrikeRatio,
                                    ContractInfo.InitialS * ContractInfo.Rebate,
                                    ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio, ContractInfo,
                                    GreeksType)),
        u'敲入看涨自动敲出赎回新起息': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.KnockInAutoCallNewEntry(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                       Dividend, ContractInfo.InitialS * ContractInfo.HstrikeRatio,
                                       ContractInfo.InitialS * ContractInfo.Rebate,
                                       ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio, ContractInfo,
                                       GreeksType)),
        u'逐步调整看涨自动敲出赎回': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.AutoCallStepChangeEntry(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol, Rf,
                                       Dividend, ContractInfo.InitialS * ContractInfo.Rebate3,
                                       ContractInfo.InitialS * ContractInfo.Rebate,
                                       ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio, ContractInfo,
                                       GreeksType)),
        u'逐步调整看涨自动敲出赎回新起息': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.AutoCallStepChangeNewEntry(PricingDate, ContractInfo.InitialS * ContractInfo.HtouchRatio, Maturity, Vol,
                                          Rf, Dividend,
                                          ContractInfo.InitialS * ContractInfo.Rebate3,
                                          ContractInfo.InitialS * ContractInfo.Rebate,
                                          ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio,
                                          ContractInfo, GreeksType)),
        u'看跌自动敲出赎回': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.AutoPutEntry(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf, Dividend,
                            ContractInfo.InitialS * ContractInfo.Rebate,
                            ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio, ContractInfo,
                            GreeksType)),
        u'看跌自动敲出赎回新起息': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.AutoPutNewEntry(PricingDate, ContractInfo.InitialS * ContractInfo.LtouchRatio, Maturity, Vol, Rf,
                               Dividend,
                               ContractInfo.InitialS * ContractInfo.Rebate,
                               ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio, ContractInfo,
                               GreeksType)),
        u'看涨自动敲出赎回转看涨': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.AutoCallableToCallEntry(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol,
                                       Rf, Dividend,
                                       ContractInfo.InitialS * ContractInfo.HtouchRatio,
                                       ContractInfo.InitialS * ContractInfo.Rebate,
                                       ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio, ContractInfo,
                                       GreeksType)),
        u'看涨自动敲出赎回转看涨新起息': lambda PricingDate, Maturity, Vol, Rf, Dividend, GreeksType, ContractInfo: (
            of.AutoCallableToCallNewEntry(PricingDate, ContractInfo.InitialS * ContractInfo.HstrikeRatio, Maturity, Vol,
                                          Rf, Dividend,
                                          ContractInfo.InitialS * ContractInfo.HtouchRatio,
                                          ContractInfo.InitialS * ContractInfo.Rebate,
                                          ContractInfo.InitialS * ContractInfo.Premium / ContractInfo.Pratio,
                                          ContractInfo, GreeksType))}


 # TODO 主力合约可以从wind api获取，当前计算基差盈亏存在问题，计算基差盈亏选取的基本合约是FutureCode中的，但是不一定是每笔期权的真是标的
    MainCode = {'HS300': 'IF.CFE',
                  'ZZ500': 'IC.CFE',
                  'SZ50': 'IH.CFE',
                  'AU': 'AU.SHF',
                  'AG': 'AG.SHF',
                  'JD': 'JD.DCE',
                  'AL': 'AL.SHF',
                  'SR': 'SR.CZC',
                  'SA': 'SA.CZC',
                  'CU': 'CU.SHF',
                  'ZN': 'ZN.SHF',
                  'M': 'M.DCE',
                  'J': 'J.DCE',
                  'PTA': 'TA.CZC',
                  'I': 'I.DCE',
                  'L': 'L.DCE',
                  'P': 'P.DCE',
                  'PP': 'PP.DCE',
                  'RB': 'RB.SHF',
                  'RU': 'RU.SHF',
                  'NI': 'NI.SHF',
                  'FG': 'FG.CZC',
                  'Y': 'Y.DCE',
                  'SF': 'SF.CZC',
                  'C': 'C.DCE',
                  'CF': 'CF.CZC',
                  'JM': 'JM.DCE',
                'A': 'A.DCE',
                  'ZC': 'ZC.CZC',
                  'OI': 'OI.CZC',
                  'MA': 'MA.CZC',
                'BU': 'BU.SHF',
                  'SC':'SC.INE'}

    FutureCode = {'HS300': 'IF2106.CFE',
                  'ZZ500': 'IC2106.CFE',
                  'SZ50': 'IH2106.CFE',
                  'AU': 'AU2106.SHF',
                  'AG': 'AG2106.SHF',
                  'JD': 'JD2105.DCE',
                  'AL': 'AL2106.SHF',
                  'SR': 'SR105.CZC',
                  'CU': 'CU2106.SHF',
                  'ZN': 'ZN2102.SHF',
                  'M': 'M2105.DCE',
                  'J': 'J2105.DCE',
                  'PTA': 'TA105.CZC',
                  'I': 'I2105.DCE',
                  'L': 'L2105.DCE',
                  'P': 'P2105.DCE',
                  'PP': 'PP2105.DCE',
                  'RB': 'RB2105.SHF',
                  'RU': 'RU2105.SHF',
                  'FG': 'FG105.CZC',
                  'Y': 'Y2105.DCE',
                  'SF': 'SF105.CZC',
                  'C': 'C2105.DCE',
                  'CF': 'CF105.CZC',
                  'JM': 'JM2105.DCE',
                  'ZC': 'ZC105.CZC',
                  'OI': 'OI105.CZC',
                  'SC':'SC2105.INE',
                  'AP':'AP010.CZC',
                  'HC':'HC2105.SHF',
                  'NI': 'NI2102.SHF',
                  'SA': 'SA105.CZC',
                  'MA': 'MA105.CZC',
                  'A': 'A2105.DCE',
                  'BU': 'BU2106.SHF',
                  '159915.SZ': '159915.SZ',
                  '399986.SZ': '512800.SH',
                  '000932.SH': '159928.SZ',
                  '399971.SZ': '512980.SH',
                  '399967.SZ': '512660.SH',
                  '159995.SZ': '159995.SZ',
                  '159996.SZ': '159996.SZ',
                  '159949.SZ': '159949.SZ',
                  '515000.SH': '515000.SH',
                  '515050.SH': '515050.SH',
                  '512760.SH': '512760.SH',
                  '512880.SH': '512880.SH',
                  '512480.SH': '512480.SH',
                  '515030.SH': '515030.SH',
                  '512010.SH': '512010.SH',
                  '300068.SZ': '300068.SZ',
                  '159966.SZ': '159966.SZ'
                  }

# 处理交易流水的字典
    CodeToUnderlying = {'IF': 'HS300',
                      '510300.SH':'HS300',
                      'IC':'ZZ500',
                      '510500.SH':'ZZ500',
                      'IH': 'SZ50',
                      '510050.SH':'SZ50',
                      'TA':'PTA',
                      '159915.SZ': '159915.SZ',
                      '512800.SH':'399986.SZ',
                      '159928.SZ':'000932.SH',
                      '512980.SH':'399971.SZ',
                      '512660.SH':'399967.SZ'}



    ParaTable = {'HS300': 'HS300',
                 'ZZ500': 'CSI500',
                 'SZ50': 'SSE50',
                 'AU': 'Aurum',
                 'CU': 'Cuprum',
                 'PTA': 'TA',
                 'I': 'Iron',
                 'L': 'LLDPE',
                 'P': 'Palm',
                 'RU': 'Rubber'}

    FuturePLCode = {'HS300': u'IF损益',
                    'ZZ500': u'IC损益',
                    'SZ50': u'IH损益',
                    'AU': u'AU损益',
                    'AG': u'AG损益',
                    'JD': u'JD损益',
                    'AL': u'AL损益',
                    'SR': u'SR损益',
                    'CU': u'CU损益',
                    'M': u'M损益',
                    'J': u'J损益',
                    'PTA': u'PTA损益',
                    'I': u'I损益',
                    'L': u'L损益',
                    'P': u'P损益',
                    'PP': u'PP损益',
                    'RB': u'RB损益',
                    'RU': u'RU损益',
                    'FG': u'FG损益',
                    'Y': u'Y损益',
                    'SF': u'SF损益',
                    'C': u'C损益',
                    'CF': u'CF损益',
                    'JM': u'JM损益',
                    'AP': u'AP损益',
                    'HC': u'HC损益',
                    'ZN': u'ZN损益',
                    'NI': u'NI损益',
                    'SA': u'SA损益',
                    'MA': u'MA损益',
                    'A': u'A损益',
                    'BU': u'BU损益',
                    'Basket0': u'Basket0损益',
                    'Basket1': u'Basket1损益',
                    'Basket2': u'Basket2损益',
                    '600837.SH': u'600837.SH损益',
                    '399967.SZ': u'399967.SZ损益',
                    '600030.SH': u'600030.SH损益',
                    '601669.SH': u'601669.SH损益',
                    '159915.SZ': u'159915.SZ损益',
                    '000712.SZ': u'000712.SZ损益',
                    '600338.SH': u'600338.SH损益',
                    '601238.SH': u'601238.SH损益',
                    '300083.SZ': u'300083.SZ损益',
                    '000793.SZ': u'000793.SZ损益',
                    '300383.SZ': u'300383.SZ损益',
                    '300130.SZ': u'300130.SZ损益',
                    '600337.SH': u'600337.SH损益',
                    '600638.SH': u'600638.SH损益',
                    '002422.SZ': u'002422.SZ损益',
                    '002042.SZ': u'002042.SZ损益',
                    '600252.SH': u'600252.SH损益',
                    '002427.SZ': u'002427.SZ损益',
                    '002146.SZ': u'002146.SZ损益',
                    '600179.SH': u'600179.SH损益',
                    '601288.SH': u'601288.SH损益',
                    '600406.SH': u'600406.SH损益',
                    '600022.SH': u'600022.SH损益',
                    '600703.SH': u'600703.SH损益',
                    '000932.SH': u'000932.SH损益',
                    '601666.SH': u'601666.SH损益',
                    '601990.SH': u'601990.SH损益',
                    '399971.SZ': u'399971.SZ损益',
                    '000723.SZ': u'000723.SZ损益',
                    '600874.SH': u'600874.SH损益',
                    '159938.SZ': u'159938.SZ损益',
                    '159996.SZ': u'159996.SZ损益',
                    'CorrPro1': u'CorrPro1损益',
                    '399986.SZ': u'399986.SZ损益',
                    'CorrPro2': u'CorrPro2损益',
                    '512000.SH':u'512000.SH损益',
                    '002602.SZ': u'002602.SZ损益',
                    '600388.SH': u'600388.SH损益',
                    '002415.SZ': u'002415.SZ损益',
                    '159949.SZ': u'159949.SZ损益',
                    '515000.SH': u'515000.SH损益',
                    '515050.SH': u'515050.SH损益',
                    '512760.SH': u'512760.SH损益',
                    '515030.SH': u'515030.SH损益',
                    '159995.SZ': u'159995.SZ损益',
                    '159966.SZ': u'159966.SZ损益',
                    '512010.SH': u'512010.SH损益',
                    '512880.SH': u'512880.SH损益',
                    '588080.SH': u'588080.SH损益',
                    '002065.SZ': u'002065.SZ损益',
                    '300068.SZ': u'300068.SZ损益',
                    '600346.SH': u'600346.SH损益',
                    '002714.SZ': u'002714.SZ损益',
                    '300498.SZ': u'300498.SZ损益',
                    '002493.SZ': u'002493.SZ损益',
                    '002251.SZ': u'002251.SZ损益',
                    '002180.SZ': u'002180.SZ损益',
                    '1918.HK': u'1918.HK损益',
                    '1833.HK': u'1833.HK损益',
                    '0772.HK': u'0772.HK损益',
                    '002010.SZ': u'002010.SZ损益',
                    '000703.SZ': u'000703.SZ损益',
                    '601225.SH': u'601225.SH损益',
                    '600779.SH': u'600779.SH损益',
                    '600893.SH': u'600893.SH损益',
                    '300003.SZ': u'300003.SZ损益',
                    '300059.SZ': u'300059.SZ损益',
                    '300496.SZ': u'300496.SZ损益',
                    '002497.SZ': u'002497.SZ损益',
                    '603363.SH': u'603363.SH损益',
                    '300369.SZ': u'300369.SZ损益',
                    '601881.SH': u'601881.SH损益',
                    '300142.SZ': u'300142.SZ损益',
                    '300378.SZ': u'300378.SZ损益',
                    '512480.SH': u'512480.SH损益',
                    '600745.SH': u'600745.SH损益',
                    '002673.SZ': u'002673.SZ损益',
                    'ZC': u'ZC损益',
                    'SC': u'SC损益',
                    'OI': u'OI损益'}

    Basket = {'Basket0': np.array([('SR701.CZC', 1 / 8), ('JD1701.DCE', 1 / 8), ('Y1701.DCE', 1 / 8),
                                   ('M1701.DCE', 1 / 8), ('RM701.CZC', 1 / 8), ('A1701.DCE', 1 / 8),
                                   ('CF701.CZC', 1 / 8), ('WH701.CZC', 1 / 8)]),
              'Basket1': np.array([('SR705.CZC', 1 / 4), ('M1705.DCE', 1 / 4),
                                   ('CF705.CZC', 1 / 4), ('P1705.DCE', 1 / 4)]),
              'Basket2': np.array([('601398.SH', 1 / 6), ('601939.SH', 1 / 6),
                                   ('601288.SH', 1 / 6), ('601328.SH', 1 / 6),
                                   ('601988.SH', 1 / 6), ('600036.SH', 1 / 6)])}

    RiskTableTitle = ['contractid', 'code', 'RiskMonteValue1', 'RiskMonteValue2',
                      'RiskMontePL', 'RiskPrice1', 'RiskPrice2', 'RiskPL',
                      'RiskDelta1', 'RiskDelta2', 'RiskDeltaPL',
                      'RiskGamma1', 'RiskGamma2', 'RiskGammaPL',
                      'RiskTheta1', 'RiskTheta2', 'RiskThetaPL',
                      'RiskVega1', 'RiskVega2', 'RiskVegaPL', 'RiskGreeksPL',
                      'RiskBiasPL', 'RiskVanna', 'RiskZomma', 'RiskVolga']
    TradeTableTitle = ['contractid', 'code', 'TradeMonteValue1', 'TradeMonteValue2',
                       'TradeMontePL', 'TradePrice1', 'TradePrice2', 'TradePL',
                       'TradeDelta1', 'TradeDelta2', 'TradeDeltaPL',
                       'TradeGamma1', 'TradeGamma2', 'TradeGammaPL',
                       'TradeTheta1', 'TradeTheta2', 'TradeThetaPL',
                       'TradeVega1', 'TradeVega2', 'TradeVegaPL', 'TradeGreeksPL',
                       'TradeBiasPL', 'TradeVanna', 'TradeZomma', 'TradeVolga']
    NewTableTitle = ['contractid', 'code', 'OptionFee', 'RiskMonteValue', 'RiskPrice',
                     'RiskDelta', 'RiskGamma', 'RiskTheta', 'RiskVega',
                     'RiskVanna', 'RiskZomma', 'RiskVolga', 'TradeMonteValue',
                     'TradePrice', 'TradeDelta', 'TradeGamma', 'TradeTheta',
                     'TradeVega', 'TradeVanna', 'TradeZomma', 'TradeVolga']

    AmericanType = [u'美式看涨', u'美式看跌']

    CorrRiskTableTitle = ['contractid', 'code', 'RiskMonteValuePre', 'RiskMonteValueCur',
                          'RiskMontePL', 'RiskPricePre', 'RiskPriceCur', 'RiskPL',
                          'RiskDeltaPL', 'RiskGammaPL', 'RiskTheta', 'RiskThetaPL',
                          'RiskVega', 'RiskVegaPL', 'RiskGreeksPL', 'RiskBiasPL']
    CorrTradeTableTitle = ['contractid', 'code', 'TradeMonteValuePre', 'TradeMonteValueCur',
                           'TradeMontePL', 'TradePricePre', 'TradePriceCur', 'TradePL',
                           'TradeDeltaPL', 'TradeGammaPL', 'TradeTheta', 'TradeThetaPL',
                           'TradeVega', 'TradeVegaPL', 'TradeGreeksPL',
                           'TradeBiasPL']
    CorrNewTableTitle = ['contractid', 'code', 'OptionFee', 'RiskMonteValue', 'RiskPrice',
                         'RiskTheta', 'RiskVega',
                         'TradeMonteValue', 'TradePrice', 'TradeTheta', 'TradeVega']

    CorrOptionDict = {u'比例价差欧式看涨': lambda x1, x2, x3, x4, x5, x6: of.Relative_Call_Entry(x1,
                                                                                         x6.InitialS * x6.HstrikeRatio,
                                                                                         x2, x3, x4, x5),
                      #x1: S_yesterday, x2: t1, x3: TempPara1, x4: TempS0,x5: 'Price', x6: singleinfo

                      u'比例价差欧式看涨鲨鱼鳍': lambda x1, x2, x3, x4, x5, x6: of.Relative_Call_EuroUOC_Entry(x1,
                                                                                                    x6.InitialS * x6.HstrikeRatio,
                                                                                                    x2,
                                                                                                    x6.InitialS * x6.HtouchRatio,
                                                                                                    x6.InitialS * x6.Rebate,
                                                                                                    x3, x4, x5)}
    CorrProduct = {'CorrPro1': np.array([('AG1812.SHF', 'AG'), ('AU1812.SHF', 'AU')]),
                   'CorrPro2': np.array([('000905.SH', 'ZZ500'), ('000016.SH', 'SZ50')])}



if __name__=='__main__':
    test = ReportInfo()
    # print(Get_Discounted_Ratio('HS300', HedgeCode['HS300']))
    import pandas as pd
    TotalContracts = pd.read_excel('D:\myprog\Contracts_and_trade\MyTotalContracts.xlsx')
    NeedContracts = TotalContracts[(TotalContracts['ContractId']=='SWHY-SJS-20009')]
    from datetime import datetime
    from Pricing.optionfunction import mytdate
    print(test.OptionDict[NeedContracts.iloc[0].OptionType](3900, 0.5, 0.2216, 0.015, 0.15, 'Price', NeedContracts.iloc[0]))
    # print(test.OptionDict['单向鲨鱼鳍看跌'](5406.24,0.04 , 0.2216, 0.015, 0.15, 'Price', NeedContracts.iloc[0]))
    # Report = test.OptionDict['看涨自动敲出赎回'](datetime(2019,12,23), , 0.15, 0, 0, 'Delta', ContractInfo)
    # test.OptionDict[u'指数增强型收益互换']
