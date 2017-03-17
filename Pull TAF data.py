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
import re
import numpy as np


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


tafline = '' 
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



#####################################
# TAF CONVERSION (UNICODE TO STRING)
stringtaf = ''
ss = re.split(r'\s+',totaltaf)
for x in ss:
    stringtaf += (x + ' ')
stringtaf = stringtaf.encode('utf-8')
#listtaf = list(stringtaf.split(' '))


#####################################
# PREP STRINGTAF
stringtaf = stringtaf.replace('BECMG','\nBECMG')
stringtaf = stringtaf.replace('TEMPO','\nTEMPO')
stringtaf = stringtaf.replace('FM','\nFM')
#print(stringtaf)


#####################################
# SPLIT TAF INTO LINES
linetaf = re.split(r'\n',stringtaf)


#####################################
# CYCLE THROUGH TAF

# CREATE TIME ARRAY
froms=[]
becomings=[]
tempos=[]
init = []
for i,line in enumerate(linetaf):
    
    #time groups
    
    if re.findall(r'(?<=FM)\d{6}',line):    
        froms += re.findall(r'(?<=FM)\d{6}',line)
    else:
        froms += [0]
    if re.findall(r'(?<=BECMG \d{4}/)\d{4}',line):
        becomings += re.findall(r'(?<=BECMG \d{4}/)\d{4}',line)
    else:
        becomings += [0]
    if re.findall(r'(?<=TEMPO \d{4}/)\d{4}',line):
        tempos += re.findall(r'(?<=TEMPO \d{4}/)\d{4}',line)
    else:
        tempos += [0]
    if i == 0:
        init += re.findall(r'\d{4}(?=/\d{4})',line)
    else:
        init += [0]
      
becomings = np.array(becomings, dtype = np.int64)
tempos = np.array(tempos, dtype = np.int64)
init = np.array(init, dtype = np.int64)
froms = np.array(froms, dtype = np.int64)
time = becomings + tempos + init + froms



#for i,line in enumerate(linetaf):    
#    if linetaf.endswith('KT'):                       #check wind groups
#        if linetaf.startswith('WS'):                   #pull out wind shear
#            WSheight = 100*int(linetaf[2:5])
#            WSdir = int(linetaf[6:9])
#            WSspeed = int(linetaf[9:11])
#            print()
#            print('HAZARD: WIND SHEAR CONDITIONS- {0}FT FROM {1} DEGREES AT {2}KT'.format(WSheight,WSdir,WSspeed))
#        else:
#            wind = int(y[-4:-2])                 #check wind thresholds
#            if wind >= 25 and wind < 35:
#                print('From {0}/{1}Z'.format(date, time))
#                print('HAZARD: SFC WIND 25-34KT')
#            elif wind >= 35 and wind < 50:
#                print('From {0}/{1}Z'.format(date, time))
#                print('HAZARD: SFC WIND 35-49KT')
#            elif wind >= 50:
#                print('From {0}/{1}Z'.format(date, time))
#                print('HAZARD: SFC WIND >50KT')
        


#for y in listtaf:                                #loop through listtaf

#    

#print(linetaf)
#print(totalicaoname) 
print(totaltaf)                                     #print for peace of mind
#print(type(totaltaf))
#print(stringtaf)
#print(type(stringtaf))
#print(listtaf)

#print(init)    
#print(froms)
#print(becomings)
#print(tempos)
print(time)