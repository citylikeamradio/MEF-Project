# -*- coding: utf-8 -*-
"""
Created on Fri Mar 10 21:23:42 2017

This program writes a TAF (requested by user) pulled from ADDS line by line 
into a specified excel cell.

@author: Robert Capella
"""
import openpyxl                           #openpxl for python excel connection
import requests                           #requests for web crawling 
import bs4                                #beautiful soup for html crawling

icao = raw_input('Enter an ICAO: ')           #asks user for icao - todo: pull from spreadsheet error check icao len/avail
tafurl = 'https://www.aviationweather.gov/taf/data?ids=' + icao +'&format=raw&metars=off&layout=off' #builds url from user input 
icaonameurl = 'http://www.airnav.com/airport/' + icao

gettaf = requests.get(tafurl)                    #saves website data
geticaoname = requests.get(icaonameurl)
gettaf.raise_for_status()                        #error checks the site pull
geticaoname.raise_for_status()  
taf = bs4.BeautifulSoup(gettaf.text,'lxml')      #saves the html
icaoname = bs4.BeautifulSoup(geticaoname.text,'lxml')
extractedtaf = taf.body.code                     #pulls the taf text within <code>
extractedicaoname = icaoname.title


q = str(extractedicaoname)
q2 = q.find('-')
q3 = q.find('/')
totalicaoname = (q[(q2+2):(q3-1)])

tafline = ' ' 
totaltaf = ''
for tafline in extractedtaf.strings:          #breaks into strings & saves formatted taf to variable
    totaltaf += (tafline + '''
          ''')                                #weird.. todo: fix


n='testbed.xlsx'                                     #variable for excel file name - exists in source folder
excelfile = openpyxl.load_workbook(n)                #load excel writer
sheet = excelfile.get_sheet_by_name('Sheet1')        #opens active work bed sheet    
sheet['A1'] = (totalicaoname)  
sheet['A2'] = (totaltaf)                             #writes taf variable to specific shell in sheet
excelfile.save(n)                                    #saves the sheet


print(totalicaoname)                                      #print for peace of mind
print(totaltaf)