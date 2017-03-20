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

#####################################
# RETRIEVE ICAO DESCRIPTION ############# UPDATE

q = str(extractedicaoname)
q2 = q.find('-')
q3 = q.find('/')
totalicaoname = (q[(q2+2):(q3-1)])

#####################################
# TAF IN UNICODE ############# UPDATE

tafline = '' 
totaltaf = ''
for tafline in extractedtaf.strings:          #breaks into strings & saves formatted taf to variable
    totaltaf += (tafline + '''
          ''')                                #weird.. todo: fix


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
temposstart = []
init = []
linesoftaf = 0
for i,line in enumerate(linetaf):
    
    #time groups
    linesoftaf += 1

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
    if re.findall(r'(?<=TEMPO )\d{4}',line):
        temposstart += re.findall(r'(?<=TEMPO )\d{4}',line)
    else:
        temposstart += [0]
    if i == 0:
        init += re.findall(r'\d{4}(?=/\d{4})',line)
    else:
        init += [0]
      
becomings = np.array(becomings, dtype = np.int64)
tempos = np.array(tempos, dtype = np.int64)
temposstart = np.array(temposstart, dtype = np.int64)
init = np.array(init, dtype = np.int64)
froms = np.array(froms, dtype = np.int64)
time = becomings + tempos + init + froms


#####################################
# CYCLE THROUGH TAF

timeend = 'End of TAF Period'
for i, line in enumerate(linetaf):
    splitline = re.split(r'\s+', line)

    if (time[i] == tempos[i]):
       timestart = temposstart[i]
       timeend = tempos[i]
    else:
        timeend = 'End of TAF Period'
        while i != (linesoftaf-1):
            timeend = time[i+1]
            break
        timestart = time[i]
    
# WIND GROUPS

    for j,group in enumerate(splitline):
        if group.endswith('KT'):                       #check wind groups
            if group.startswith('WS'):                   #pull out wind shear
                WSheight = 100*int(group[2:5])
                WSdir = int(group[6:9])
                WSspeed = int(group[9:11])
                print(time[i])
                print('HAZARD: WIND SHEAR CONDITIONS- {0}FT FROM {1} DEGREES AT {2}KT'.format(WSheight,WSdir,WSspeed))
            else:
                wind = int(group[-4:-2])                 #check wind thresholds
                if wind >= 25 and wind < 35:
                    print('From {0} to {1}: HAZARD: SFC WIND 25-34KT'.format(timestart,timeend))
                elif wind >= 35 and wind < 50:
                    print('From {0} to {1}: HAZARD: SFC WIND 35-49KT'.format(timestart,timeend))
                elif wind >= 50:
                    print('From {0} to {1}: HAZARD: SFC WIND >50KT'.format(timestart,timeend))
                    
#VIS GROUP - M and SM  
    vis = 9999              
    for j, group in enumerate(splitline):
         if len(group) == 4 and re.findall(r'\d{4}',group) and j<6:
            vis = int(group)
            if vis <=8000 and vis >4800:
                print('From {0} to {1}: HAZARD: 3SM < VIS < 5SM'.format(timestart,timeend))
            elif vis<=4800 and vis > 1600:
                print('From {0} to {1}: HAZARD: 1SM < VIS < 3SM'.format(timestart,timeend))
            elif vis <=1600:
                print('From {0} to {1}: HAZARD: VIS < 1SM'.format(timestart,timeend))
         elif (len(group) == 3 or len(group) == 5 or len(group) == 6) and re.findall(r'SM', group):
             vis = group
             if splitline[j-1] == '1':
                 vis = '1 {0}'.format(vis)
             elif splitline[j-1] == '2':
                 vis = '2 {0}'.format(vis)
             ifr = dict([('1SM', 0),('1 1/8SM', 1), ('1 1/4SM', 2), ('1 3/8SM', 3), ('1 1/2SM', 4), ('1 5/8SM', 5), ('1 3/4SM', 6), ('1 7/8SM', 7), ('2SM', 8), ('2 1/4SM', 9), ('2 1/2', 10), ('2 3/4SM', 11)])
             lifr = dict([('0SM', 90),('1/16SM', 91), ('1/8SM', 92), ('3/16SM', 93), ('1/4SM', 94), ('5/16SM', 95), ('3/8SM', 96), ('1/2SM', 97), ('5/8SM', 98), ('3/4SM', 99), ('7/8', 910), ('1SM', 911)])
             if vis == '5SM' or vis == '4SM' or vis == '3SM':
                print('From {0} to {1}: HAZARD: 3SM < VIS < 5SM'.format(timestart,timeend))
             elif vis in ifr:
                print('From {0} to {1}: HAZARD: 1SM < VIS < 3SM'.format(timestart,timeend))
             elif vis in lifr:
                print('From {0} to {1}: HAZARD: VIS < 1SM'.format(timestart,timeend))
                 
             
    
#####################################
# WRITE TO EXCEL FILE ############# UPDATE                
                    
n='testbed.xlsx'                                     #variable for excel file name - exists in source folder
excelfile = openpyxl.load_workbook(n)                #load excel writer
sheet = excelfile.get_sheet_by_name('Sheet1')        #opens active work bed sheet    
sheet['A1'] = (totalicaoname)  
sheet['A2'] = (totaltaf)                             #writes taf variable to specific shell in sheet
excelfile.save(n)                                    #saves the sheet

#####################################
# BLOCK OF PRINT STATEMENTS 

#print(linetaf)
print(totalicaoname) 
print(totaltaf)                                     #USE TO NOT GO CRAZY
#print(type(totaltaf))
#print(stringtaf)
#print(type(stringtaf))
#print(listtaf)
#print(init)    
#print(froms)
#print(becomings)
#print(tempos)
#print(temposstart)
#print(time)
#print linesoftaf