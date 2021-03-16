# -*- coding: utf-8 -*-
"""
Created on Fri Mar 12 19:05:41 2021

@author: Kunal Sinha
"""
#importing the libraries
import openpyxl as xl

#loading the data files i.e. workbooks
filename = r"C:/Users/Dell/Desktop/Datafolder/Data1.xlsx"
wb = xl.load_workbook(filename)

#loading the sheet
ws = wb.sheetnames
print(ws)

def process_spredsheet(filename):
    for sht in ws:
        print(sht)
      
        sheet = wb[sht]
        
        def process(sheet):
            #automating the formula for excel using for loop
            for i in range(23):
                sheet['C'+str(4+i)] = '=B'+str(4+i)+'/10000'
                
            for i in range (23):
                sheet['D'+ str(5+i)] = '=(' + str(i) + r'/10000)'
               
            
            sheet['E4'] = '100'            
            for i in range(23):
                sheet['E'+str(5+i)] = '=E'+str(4+i)+'*(1+D'+str(5+i)+')'
                
            wb.save(filename)
            
        process(sheet)


process_spredsheet(filename)


