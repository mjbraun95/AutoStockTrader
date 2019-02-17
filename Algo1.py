from __future__ import division
from zipline.pipeline.factors import AverageDollarVolume, CustomFactor, Returns
from zipline.pipeline.filters import StaticAssets
import datetime
import math
import numpy as np
from operator import itemgetter
import pandas as pd
from scipy import stats
import talib


class SecurityInList(CustomFactor):
    inputs = []
    window_length = 1
    securities = []
    def compute(self, today, assets, out):
        out[:] = np.in1d(assets, self.securities)

def initialize(context):
    """
    Called once at the start of the algorithm.
    """
    context.active = [sid(41968), sid(37514), sid(45705), sid(19658), sid(28073)]
    context.absolute = sid(40670)
    context.outmkt = sid(23911) # SHY right now, will re make this with vixy

    context.assets = set(context.active + [context.absolute, context.outmkt])    
    context.currently_holding = []

    context.lookback = 20
    set_benchmark(sid(8554)) # SPY ## (sid(40107)) # VOO ## sid(21513)) # IVV ##
    set_long_only()
    context.enable_short_stock = False

    set_commission(commission.PerShare(cost=0, min_trade_cost=0))

    context.leverage = 1.0
    context.allocation = pd.Series([0.0]*len(context.assets), index = context.assets)


    schedule_function(my_mom, 
                      date_rules.week_end(),
                      time_rules.market_close(hours=2, minutes=30))
    #schedule_function(rebalance, 
    #                 date_rules.week_end(),
    #                 time_rules.market_close(hours=2, minutes=30))   


def my_mom(context, data):
    best_mom = []

    histabs = pd.DataFrame([data.history(context.absolute, 'price', context.lookback+1, '1d')])
    absmom = (histabs.iloc[-1]/histabs.iloc[0]) - 1.0
    absmom = absmom.dropna()
    log.info(absmom.head())
    positive_absmom = histabs[-5:].mean() > histabs.mean()

    if positive_absmom is True:
        if data.can_trade(context.absolute):
            best_mom.append(context.absolute)
            print(best_mom)


    assets = context.active
    hist = pd.DataFrame([data.history(assets, 'price', context.lookback+1, '1d')])
    relmom = (hist.iloc[-1]/hist.iloc[0]) - 1.0
    relmom = relmom.dropna()
    top_rel = relmom.sort_values().index[-5:]
    top_rel_val = relmom.sort_values()[-5:]
    top_rel_val = sorted(assets, key = itemgetter, reverse = True)
    log.info(top_rel_val[0])

    positive_relmom = hist[-5:].mean() > hist.mean()

    if positive_relmom is True:
        if top_rel_val[0] not in best_mom:
            if data.can_trade(top_rel_val[0]):
                best_mom.append(top_rel_val[0])
                print(best_mom)

    N = len(best_mom)

    if N < 1:
        if data.can_trade(context.outmkt):
            order_target_percent(context.outmkt, 1.0)
    else:
        best_mom = sorted(best_mom, key = itemgetter(1), reverse = True)
        best_mom = best_mom[0:1]
        print (best_mom[0][0].symbol)


    best_mom = sorted(best_mom, key = itemgetter(1), reverse = True)
    for stock in context.currently_holding:
        order_target_percent(stock, 0)
        context.currently_holding = []
        context.currently_holding.append(best_mom)
        if len(context.currently_holding) == 1:
            order_target_percent(stock, 1.0)