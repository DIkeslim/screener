import datetime as dt
import pandas as pd
from pandas_datareader import data as pdr
import yfinance as yf

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os
from pandas import ExcelWriter

yf.pdr_override()
start =dt.datetime(2017,12,1)
now = dt.datetime.now()

filePath=r"C:\richardstocks\RichardStocks.xlsx"

stocklist = pd.read_excel(filePath)
stocklist=stocklist.head()

exportList = pd.DataFrame(columns=['Stock', "RS_Rating", "50 Day MA", "150 Day Ma", "200 Day MA", "52 Week Low", "52 week High"])


for i in stocklist.index:
    stock=str(stocklist["Symbol"][i])
    Rs_Rating=stocklist["RS Rating"][i]

    try:
        df = pdr.get_data_yahoo(stock, start, now)

        smaUsed=[50,100,200]
        for x in smaUsed:
            sma=x
            df["SMA_"+str(sma)]=round(df.iloc[:4].rolling(window=sma).mean(),2)

        currentClose=df["Adj Close"][-1]
        moving_average_50=df["SMA_50"][-1]
        moving_average_150=df["SMA_150"][-1]
        moving_average_200=df["SMA_200"][-1]
        low_of_52week=min(df["Adj Close"][-260:])
        high_of_52week=max(df["Adj Close"][-260:])

        try:
            moving_average_200_20past=df["SMA_200"][-20]
        except Exception:
            moving_average_200_20past=0

        print("Checking "+stock+"...")

        #condition 1: Current price > 150 SMA and > 200 SMA
        if(currentClose>moving_average_150 and currentClose>moving_average_200):
            cond_1=True
        else:
            cond_1=False
        #condition 2: 150 SMA and > 200 SMA
        if(moving_average_150>moving_average_200):
            cond_2 = True
        else:
            cond_2 = False
        #condition 3: 200 SMA trending for at least 1 month (ideally 4-5 months)
        if (moving_average_200 > moving_average_200_20past):
            cond_3= True
        else:
            cond_3= False
        #condition 4: 50 SMA>150 and 50 SMA> 200 SMA
        if(moving_average_50> moving_average_150 and moving_average_50> moving_average_200):
            cond_4 = True
        else:
            cond_4 = False
        #condition 5: Current price > 50 SMA
        if(currentClose>=moving_average_50):
            cond_5 = True
        else:
            cond_5 = False

        #condition 6: Current price is at least 30% above 52 week low (many of the best are up 100-300)
        if(currentClose>=(1.3*low_of_52week)):
            cond_6 = True
        else:
            cond_6 = False
        #condition 7: Current price is within 25% of 52 weeks high
        if (currentClose >= (75 * high_of_52week)):
            cond_7 = True
        else:
            cond_7 = False

        #condition 8 : IBD RS rating >70 and the higher the better
        if(Rs_Rating>70):
            cond_8 = True
        else:
            cond_8 = False

        if(cond_1 and cond_2 and cond_3 and cond_4 and cond_5 and cond_6 and cond_7 and cond_8):
            exportList = exportList.append({'Stock': stock, "RS_Rating": Rs_Rating, "50 Day MA": moving_average_50, "150 Day MA": moving_average_150,
                                            "200 Day MA": moving_average_200, "52 Week Low": low_of_52week, "52 Week High": high_of_52week})
    except Exception:
        print("No data on" +stock)

print(exportList)





