###Additional text.###

import Tkinter
import tkFileDialog
import os
import csv
import time
import pandas as pd
from pandas import Series, DataFrame
import matplotlib.pyplot as plt
import matplotlib
import pylab
import numpy
from pylab import figure, axes, title, show
from jinja2 import Environment, FileSystemLoader

###Tkinter window handler###
root = Tkinter.Tk()
root.withdraw()
###Jinja2 calling file and template###
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template('/Report/callvolumereport.html')
### Additional Text###
print 'You will need to down load the CDR.csv file from CUCM'
print 'Navigate to Cisco Unified Serviceability >>>'
print 'Tools >>> CDR Analysis and Reporting. Select Dates and download!'
print ' '
print '---------------------------****---------------------------------'
print ' '
###Defineing variables and lists that will be needed###
###There are a few other variables that will still be defined###
month = raw_input('Enter month or time frame this report is for: ')
department = raw_input('What department is this for? ')
print ' '
print '---------------------------****---------------------------------'
print ' '
phonenumbers = []
image_list_per_day = []
image_list_per_hour = []
users_dict = {}
###This will call window open dialog to load dcr record###
ofile = tkFileDialog.askopenfilename()
root.destroy()
tempdata = 'C:/Users/nlmorgan/Documents/Python/cdr/tempdata.csv'
f1 = open(ofile, 'r')
f2 = open(tempdata, 'w')
###Ask the user for a list of phone numbers, max limit set to find 20 numbers###
print '+You can look for 5 digit extentions for DCSD numbers'
print '+Or you can search for an outside number, enter without -'
print '+Example: 17204331212'
print '+You can enter up to 20 phone numbers.'
print '!!! Type "done" when you are ready to run the report !!!'
print ' '
print '---------------------------****---------------------------------'
print ' '
maxLengthList = 20
while len(phonenumbers)<maxLengthList:
    number = raw_input('Enter phone number/s to look for: ')
    if number == 'done':
        break
    phonenumbers.append(number)
    print 'The following numbers will be in your report:'+(', ' .join(phonenumbers))
###Data mining to run the report###
with open(tempdata, 'r+') as f:
    writer = csv.writer(f, lineterminator ='\n')
    for i in phonenumbers:
        with open(ofile, 'r') as csvfile:
            reader = csv.DictReader(csvfile, delimiter = ',')
            for row in reader:
                if row['callingPartyNumber'] == i or row['finalCalledPartyNumber'] == i:
                    writer.writerow([time.ctime(float(row['dateTimeOrigination'])),
                                     time.ctime(float(row['dateTimeConnect'])),
                                     time.ctime(float(row['dateTimeDisconnect'])),
                                     row['callingPartyNumber'],
                                     row['finalCalledPartyNumber']])
                    users_dict.update(row)

###Reading the CDR and set up the data we want to see###
callvolume = pd.read_csv(tempdata, names=['dateTimeOrigination','dateTimeConnect',
                                          'dateTimeDisconnect','Calling','Called'],
                         dtype=object)
callvolume.dateTimeOrigination = pd.to_datetime(callvolume.dateTimeOrigination)
rcvd = DataFrame({'Date':callvolume['dateTimeOrigination'],
                  'Called':callvolume['Called']})
rcvd_sumary_data = rcvd.describe(include='all')
grouped2 = rcvd['Called'].groupby(rcvd['Date'].dt.date)
grouped2_data = grouped2.describe(include='all')
grouped3 = dict(list(rcvd['Date'].dt.date.groupby(rcvd['Called'])))
grouped4 = dict(list(rcvd['Date'].dt.hour.groupby(rcvd['Called'])))

###Generate graphs for the report###
###Graph of phone numbers by date###
for i in phonenumbers:
    if i in list(rcvd['Called']):
        print i+' was called:'
        globals()['daydata%s' %i] = grouped3[i].groupby(rcvd['Date'].dt.date).count()
        dayfig = plt.figure(figsize=(9,5))
        globals()['daydata%s' %i].plot(kind='bar', color='#7C7D7D')
        #dayfig.Color = [75 110 165]
        plt.title('Number of Times %s was called in\n' %i + month, fontsize=20)
        plt.xlabel('', fontsize=12)
        plt.xticks(rotation=45, horizontalalignment='right')
        plt.tight_layout(pad=0.4, w_pad=0.5, h_pad=1.5)
        plt.ylabel('# of Calls', fontsize=12)
        plt.subplots_adjust(bottom=0.25)
        dayfig.savefig('daygraph%s.png' %i)
        image_list_per_day.append("""<img style="width: 450px; height: 350px;" src="daygraph%s.png" alt="">"""
                                  %i)
        dayfig.clear()
###Graph of phone numbers by hour###
for x in phonenumbers:
    if x in list(rcvd['Called']):
        print x+' was called:'
        globals()['hourdata%s' %x] = grouped4[x].groupby(rcvd['Date'].dt.hour).count()
        hourfig = plt.figure(figsize=(9,5))
        globals()['hourdata%s' %x].plot(kind='bar', color='#7291A5')
        #hourfig.Color = [60 90 145]
        plt.title('Number of Times %s \n was called by hour for ' %x + month, fontsize=18)
        plt.xlabel('24hr clock' , fontsize=12)
        plt.xticks(rotation=0, horizontalalignment='center')
        plt.subplots_adjust(bottom=0.15)
        plt.ylabel('# of Calls', fontsize=12)
        plt.tight_layout(pad=0.4, w_pad=0.5, h_pad=1.5)
        hourfig.savefig('hourgraph%s.png' %x)
        image_list_per_hour.append("""<img style="width: 450px; height: 350px;" src="hourgraph%s.png" alt="">"""
                                  %x)
        dayfig.clear()
###Summary information###
sumfig = plt.figure(figsize=(9,5))
grouped2.count().plot(kind='bar', color='#093F64')
plt.title('Summary of calls in '+month, fontsize=20)
plt.xlabel('', fontsize=12)
plt.xticks(rotation=45, horizontalalignment='right')
plt.ylabel('# of Calls', fontsize=12)
plt.subplots_adjust(bottom=0.25)
plt.tight_layout(pad=0.4, w_pad=0.5, h_pad=1.5)
sumfig.savefig('sumfig.png')
sumfig.clear()
html_sumfig = """<img style="width: 450px; height: 350px" src="sumfig.png" alt="">"""

###Collecting variables applying and writing them to the template###
template_vars = {"month":month, "department":department, "phone_numbers":(', ' .join(phonenumbers)),
                 "s_variable":rcvd_sumary_data.to_html(border=0), "s_graph":html_sumfig,
                 "describe_variable":grouped2_data.to_html(border=0),
                 "image_list_per_day":(' ' .join(image_list_per_day)),
                 "image_list_per_hour":(' ' .join(image_list_per_hour))}
html_out = template.render(template_vars)
with open('report.html', 'wb')as fh:
    fh.write(html_out)
os.startfile('report.html')
f1.close()
f2.close()
os.remove(tempdata)
SystemExit()

                 

