# -*- coding: utf-8 -*-
"""
Created on Fri May 14 19:08:10 2021

@author: maxim
"""
import pandas as pd
import yfinance as yf
import numpy as np
import xlsxwriter as xl
from tkinter import filedialog
import tkinter as tk

root = tk.Tk()
root.title("Maximilian's Dividend Discount Model")

tlbl = tk.Label(root, text='Ticker')
tlbl.grid(column=0, row=0)
Name = tk.Entry(root,width=10)
Name.grid(column=1, row=0)

caplbl = tk.Label(root, text='CAPM or Expected Return')
caplbl.grid(column=0, row=1)
CAPM = tk.Entry(root,width=10)
CAPM.grid(column=1, row=1)

stlbl = tk.Label(root, text='Short Term Dividend Growth Rate')
stlbl.grid(column=0, row=2)
STg = tk.Entry(root,width=10)
STg.grid(column=1, row=2)

tvlbl = tk.Label(root, text='Terminal Value Dividend Growth Rate')
tvlbl.grid(column=0, row=3)
TVg = tk.Entry(root,width=10)
TVg.grid(column=1, row=3)

drlbl = tk.Label(root, text='Discount Rate')
drlbl.grid(column=0, row=4)
DiscountRate = tk.Entry(root,width=10)
DiscountRate.grid(column=1, row=4)

fplbl = tk.Label(root, text='Forecasting Periods')
fplbl.grid(column=0, row=5)
FP = tk.Entry(root,width=10)
FP.grid(column=1, row=5)


def DDM():
    try:  
        #getting ticker and calling yahoo finance function
        tick = Name.get()
        Ticker = yf.Ticker(tick)
        
        #Calculating Beta
        Ticker1 = yf.Ticker('^GSPC')
        Company = pd.DataFrame(Ticker.history(period = "5y",interval = "1mo"))
        Comp = Company['Close']
        SP = pd.DataFrame(Ticker1.history(period = "5y",interval = "1mo")['Close'])
        data1 = pd.concat([Comp, SP],axis = 1)
        sec_returns = np.log( data1 / data1.shift(1) )
        cov = sec_returns.cov() * 250
        cov_with_market = cov.iloc[0,1]
        market_var = sec_returns.iloc[:,1].var() * 250
        beta = cov_with_market / market_var
        
        # Get names of indexes for which columns have a value of 0
        indexNames = Company[ Company['Dividends'] == 0.00 ].index
        # Delete these row indexes from dataFrame
        Company.drop(indexNames , inplace=True)
        
        #list of dividend payments only
        dividends = Company['Dividends']
        dividends = dividends.reset_index()
        
        #Creating a new column with only the year instead of full date
        years = []
        dividends = pd.DataFrame(dividends)
        for d in dividends["Date"]:
            years.append(str(d)[:4])
            
        #finding the last paid dividned value
        last_div = dividends['Dividends'][len(dividends)-1]
        
        #making the new year only column the main column for date instead of full datetime
        dividends['Year'] = years
        dividends.drop(columns="Date")
        
        #finding how many dividends have been paid each year
        dividends = dividends[['Year','Dividends']].set_index('Year')
        dups_divs = dividends.pivot_table(index=['Year'], aggfunc='size')
        
        #making a value for the current year that will be their full dividend payment for the year then replacing that value in the dataframe
        yearly = dividends.sum(level = 'Year')
        new_div = last_div * 4
        yearly['Dividends'] = yearly['Dividends'].replace(dups_divs[-1]*last_div,new_div)
        df = yearly.reset_index()
        dividends_list = df.values.tolist()
        
        #finding path
        filename2 = filedialog.asksaveasfilename(
                    defaultextension='.xlsx', filetypes=[("Excel file", '*.xlsx')],
                    title="Choose filename")
        path = str(filename2)
        
        # Create a workbook and add a worksheet.
        workbook = xl.Workbook(path)
        worksheet1 = workbook.add_worksheet('Company Name')
        worksheet2 = workbook.add_worksheet('DDM')
        
        # Start from the first cell. Rows and columns are zero indexed.
        row1 = 1
        col1 = 0
        
        # Iterate over the data and write it out row by row.
        for date, div in (dividends_list):
            try:
                worksheet1.write(0,0,tick)
                worksheet1.write(0,1,'Beta = '+ str(beta))
                worksheet1.write(row1 + 1, col1, date)
                worksheet1.write(row1 + 1, col1 + 1, div)
                worksheet1.write(row1 + 2, col1 + 2, '=B'+ str(3+row1) + '/B' + str(2+ row1)+ '-1')
                row1 += 1
            except:
                pass
            
        # Write a total using a formula.
        worksheet2.write(0, 0, 'Dividned Discount Method')
        worksheet2.write(2, 0, 'CAPM/r')
        worksheet2.write(2, 1, CAPM.get())
        worksheet2.write(3, 0, 'Short Term Dividend Growth')
        worksheet2.write(3, 1, STg.get())
        worksheet2.write(4, 0, 'Terminal Value Dividend Growth')
        worksheet2.write(4, 1, TVg.get())
        worksheet2.write(5, 0, 'Discount Rate')
        worksheet2.write(5, 1, DiscountRate.get())
        worksheet2.write(7, 0, "This Years Dividend")
        worksheet2.write(8, 0, "='Company Name'!B" + str(row1+1))
        
        
        #first writing how many dicount periods, then next row is a forecast, then next row its dicounting forecast
        e = 0
        periods = int(FP.get())
        for n in range(periods):
            worksheet2.write(7, n+1, str(n+1))
            worksheet2.write(8, n+1, "=A9*(1+B4)^" + chr(ord('@')+2+n)+ "8")
            worksheet2.write(9, n+1, "=(A9*((1+B4)^" + chr(ord('@')+2+n)+ "8" + "))/((1+B6)^"+chr(ord('@')+2+n)+ "8"+ ")")
            e+=1
            
        #Building Terminal Value
        worksheet2.write(7, e+1, str(e+1))
        worksheet2.write(8,e+1, "=(A9*(1+B4)^" + chr(ord('@')+2+e)+ "8" +")/(B3-B5)")
        worksheet2.write(9,e+1, "=((A9*(1+B4)^" + chr(ord('@')+2+e)+ "8" +")/(B3-B5))" + "/((1+B6)^"+chr(ord('@')+1+e)+ "8"+ ")")
        
        #Adding all discounted values
        worksheet2.write(11,0, "Present Value of Dividends")
        worksheet2.write(11,1, "=SUM(B10:" + chr(ord('@')+2+e) + "10)")
        workbook.close()
        
        tk.messagebox.showinfo('Complete','Finished')
    
    except:
        tk.messagebox.showinfo('Error','Something went wrong double check numbers')


DDMbtn = tk.Button(root, text="RUN", command=DDM)
DDMbtn.grid(column=2, row=7)

root.mainloop()