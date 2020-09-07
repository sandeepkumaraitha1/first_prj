# -*- coding: utf-8 -*-
"""
Created on Tue Jun 23 16:27:40 2020

@author: SANDEEP
"""
"""import xlrd

loc=("E:/python/readfile/Report.xlsx")
print("reading")
wb=xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
print(sheet.nrows,sheet.name)"""
"""
to show data in pie diagram 

import pandas as pd 
import matplotlib.pyplot as plt
loc=("E:/python/readfile/Report.xlsx")
df = pd.read_excel(loc,"RAN_Report")
colors = ['yellowgreen', 'gold', 'lightskyblue']
df['Parent Design Type'].value_counts().plot(kind='pie',title='Design Types',colors=colors)
plt.show()
"""
"""
EXCEL READING 
"""

import pandas as pd 
loc = ("E:/python/readfile/Report.xlsx")
loc1 = ("E:/python/readfile")
df = pd.read_excel(loc,"RAN_Report")
col=df.columns
for i in col:
    #print("columns",i)
    if i == "SR No":
        #print("i",df['SR No']==1)
        print(df.loc[2])
        d1=df.loc[2]
        d1.to_excel("o1.xlsx")
#import datetime as dt
#print(min(10,12))
#try:
 #   print(["1","2"])
##   print("error in thing")
    


"""
EXCEL WRITTING 

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np

df = pd.DataFrame({'a':[1,3,5,7,4,5,6,4,7,8,9],
                   'b':[3,5,6,2,4,6,7,8,7,8,9]})

writer = ExcelWriter('Pandas-Example2.xlsx')
df.to_excel(writer,'Sheet1',index=False)
writer.save()
"""