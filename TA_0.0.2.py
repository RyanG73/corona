import git
import pandas as pd
import numpy as np
import glob
import datetime as dt
import seaborn as sns
import matplotlib.pyplot as plt
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from sklearn import preprocessing
import matplotlib.dates as mdates

start_time = time.time()
clock = time.strftime('%x')
today = time.strftime('%x')
yesterday = (time.time() - 86400)
yesterday = dt.datetime.fromtimestamp(yesterday)
yesterday = yesterday.strftime('%x')
clock = clock.replace("/","_")

repo = git.Repo('/Users/ryangerda/PycharmProjects/corona/COVID-19')
repo.remotes.origin.pull()

path = r'/Users/ryangerda/PycharmProjects/corona/COVID-19/csse_covid_19_data/csse_covid_19_daily_reports_us' # use your path
all_files = glob.glob(path + "/*.csv")
li = []
df = pd.concat((pd.read_csv(f) for f in all_files))

census = "/Users/ryangerda/PycharmProjects/corona/Census_2019.xlsx"
census = pd.read_excel(census)


df['Last_Update'] = pd.to_datetime(df['Last_Update']).dt.normalize()
df = df.rename(columns={'Province_State':'State','Last_Update':'Date'})
states = ['Alabama','Alaska','Arizona', 'Arkansas',
           'California', 'Colorado', 'Connecticut', 'Delaware',
           'District of Columbia','Florida', 'Georgia','Hawaii',
           'Idaho', 'Illinois', 'Indiana',
           'Iowa', 'Kansas', 'Kentucky', 'Louisiana', 'Maine', 'Maryland',
           'Massachusetts', 'Michigan', 'Minnesota', 'Mississippi',
           'Missouri', 'Montana', 'Nebraska', 'Nevada', 'New Hampshire',
           'New Jersey', 'New Mexico', 'New York', 'North Carolina',
           'North Dakota','Ohio', 'Oklahoma',
           'Oregon', 'Pennsylvania','Rhode Island',
           'South Carolina', 'South Dakota', 'Tennessee', 'Texas', 'Utah',
           'Vermont','Virginia', 'Washington',
           'West Virginia', 'Wisconsin', 'Wyoming']
#states = ['West Virginia', 'Wisconsin']
df = df[df['State'].isin(states)]
df = df.sort_values('Date')
df = df.reset_index(drop=True)
df = pd.merge(df,census[['Area',2019]],how='left',left_on='State',right_on='Area')
del df['Area']
df = df.rename(columns={2019:'Population_2019'})
df2 = pd.DataFrame()

writer = pd.ExcelWriter(('COVID_' + str(clock) + '.xlsx'), engine='xlsxwriter')
workbook = writer.book

for i in states:
    i2 = [i]
    d = df[df['State'].isin(i2)].copy()
    d['daily_newcases'] = d['Confirmed'].diff().fillna(0)
    d['daily_newcases'] = np.where(d['daily_newcases'] < 0,0,d['daily_newcases'])
    d['rolling_newcases'] = d['daily_newcases'].rolling(window=7).mean()
    d['daily_deaths'] = d['Deaths'].diff().fillna(0)
    d['daily_deaths'] = np.where(d['daily_deaths'] < 0, 0, d['daily_deaths'])
    d['rolling_deaths'] = d['daily_deaths'].rolling(window=7).mean()
    d['daily_recovered'] = d['Recovered'].diff().fillna(0)
    d['daily_recovered'] = np.where(d['daily_recovered'] < 0, 0, d['daily_recovered'])
    d['rolling_recovered'] = d['daily_recovered'].rolling(window=7).mean()
    d['daily_tested'] = d['People_Tested'].diff().fillna(0)
    d['daily_tested'] = np.where(d['daily_tested'] < 0, 0, d['daily_tested'])
    d['rolling_tested'] = d['daily_tested'].rolling(window=7).mean()
    d['daily_hospitalized'] = d['People_Hospitalized'].diff().fillna(0)
    d['daily_hospitalized'] = np.where(d['daily_hospitalized'] < 0, 0, d['daily_hospitalized'])
    d['rolling_hospitalized'] = d['daily_hospitalized'].rolling(window=7).mean()
    d['hospitalization_change'] = d['rolling_hospitalized'] / d['rolling_hospitalized'].shift(7) - 1
    d['cases_to_population'] = d['daily_newcases'] / d['Population_2019'] * 1000
    d['hospitalization_to_population'] = d['daily_hospitalized'] / d['Population_2019'] * 1000
    d['deaths_to_population'] = d['daily_deaths'] / d['Population_2019'] * 1000
    d['rolling_case_ratio'] = d['cases_to_population'].rolling(window=7).mean()
    d['rolling_hospitalization_ratio'] = d['hospitalization_to_population'].rolling(window=7).mean()
    d['rolling_death_ratio'] = d['deaths_to_population'].rolling(window=7).mean()
    d['rolling_death_index'] = (d['rolling_death_ratio'] - d['rolling_death_ratio'].min()) / (
                d['rolling_death_ratio'].max() - d['rolling_death_ratio'].min())
    d['rolling_hospitalization_index'] = (d['rolling_hospitalization_ratio'] - d['rolling_hospitalization_ratio'].min()) / (
                d['rolling_hospitalization_ratio'].max() - d['rolling_hospitalization_ratio'].min())
    d['rolling_case_index'] = (d['rolling_case_ratio'] - d['rolling_case_ratio'].min()) / (
                d['rolling_case_ratio'].max() - d['rolling_case_ratio'].min())
    d = d[(d['Date'] > '2020-04-20')]
    df2 = df2.append(d)


data_chart2 = df2[['State','Date','rolling_case_index', 'rolling_hospitalization_index','rolling_death_index']]
df2.to_excel(writer, sheet_name='Data', index=False)
data_sheet = writer.sheets['Data']
data_sheet.set_column(0,100,15)

for i in states:
    i2 = [i]
    d2 = df2[df2['State'].isin(i2)].copy()
    d3 = data_chart2[data_chart2['State'].isin(i2)].copy()
    empty = pd.DataFrame()
    empty.to_excel(writer, sheet_name=str(i))
    chart_sheet = writer.sheets[str(i)]
    chart_sheet.hide_gridlines(2)
    # CHART 1
    plt.figure(figsize=(6,6))
    plt.plot(d2['Date'],d2["rolling_tested"], color='Red', label='7-Day Rolling Average')
    plt.bar(d2['Date'],d2["daily_tested"], color='Blue', label='Daily Count')
    plt.xticks(rotation=20)
    plt.xlabel("Date")
    plt.ylabel("Number Tested")
    plt.legend(loc='upper right')
    plt.title(str('Testing - ' + i))
    filename = 'charts/' + str(i) + '_1.png'
    fig = plt.savefig(filename,bbox_inches='tight')
    plt.close('all')
    chart_sheet.insert_image('A30', filename)
    print(i)
    # CHART 2
    fig, ax = plt.subplots(figsize=(6,6))
    ax.plot(d2['Date'],d2["rolling_hospitalized"], color='Red', label='7-Day Rolling Average')
    ax.bar(d2['Date'],d2["daily_hospitalized"], color='Blue', label='Daily Count')
    plt.xticks(rotation=20)
    ax.set_xlabel("Date")
    ax.set_ylabel("Number Hospitalized")
    ax.legend(loc='upper right')
    plt.title(str('Hospitalizations - ' + i))
    filename = 'charts/' + str(i) + '_2.png'
    fig = plt.savefig(filename,bbox_inches='tight')
    plt.close('all')
    chart_sheet.insert_image('J1',filename)
    print(i)
    # CHART 3
    fig, ax = plt.subplots(figsize=(6,6))
    ax.plot(d2['Date'], d2["hospitalization_change"], color='Red')
    #axe = sns.lineplot(x=df2['Date'], y="hospitalization_change", data=df2)
    ax.axhline(0, ls='--',color='black')
    plt.xticks(rotation=20)
    ax.set_xlabel("Date")
    ax.set_ylabel("Percent Change")
    plt.title(str('7-Day Rolling Change in Hospitalization - ' + i))
    filename = 'charts/' + str(i) + '_3.png'
    fig = plt.savefig(filename,bbox_inches='tight')
    plt.close('all')
    chart_sheet.insert_image('S30',filename)
    print(i)
    # CHART 4
    fig, ax = plt.subplots(figsize=(6,6))
    ax.plot(d2['Date'],d2["rolling_newcases"], color='Red', label='7-Day Rolling Average')
    ax.bar(d2['Date'],d2["daily_newcases"], color='Blue', label='Daily Count')
    plt.xticks(rotation=20)
    ax.set_xlabel("Date")
    ax.set_ylabel("Count")
    ax.legend(loc='upper right')
    plt.title(str('New Cases - ' + i))
    filename = 'charts/' + str(i) + '_4.png'
    fig = plt.savefig(filename,bbox_inches='tight')
    plt.close('all')
    chart_sheet.insert_image('S1',filename)
    print(i)
    # CHART 5
    fig, ax = plt.subplots(figsize=(6,6))
    ax.plot(d2['Date'],d2["rolling_deaths"], color='Red', label='7-Day Rolling Average')
    ax.bar(d2['Date'],d2["daily_deaths"], color='Blue', label='Daily Count')
    plt.xticks(rotation=20)
    ax.set_xlabel("Date")
    ax.set_ylabel("Count")
    ax.legend(loc='upper right')
    plt.title(str('Deaths - ' + i))
    filename = 'charts/' + str(i) + '_5.png'
    fig = plt.savefig(filename,bbox_inches='tight')
    plt.close('all')
    chart_sheet.insert_image('A1',filename)
    print(i)
    # CHART 6
    fig, ax = plt.subplots(figsize=(6,6))
    p1 = sns.lineplot(x='Date', y='value', hue='variable',data=pd.melt(d3, ['State','Date']),legend='brief')
    ax.axhline(0.5, ls='--', color='black')
    plt.xticks(rotation=20)
    ax.set_xlabel("Date")
    ax.set_ylabel("Index")
    #ax.legend(loc='upper right')
    handles, labels = ax.get_legend_handles_labels()
    ax.legend(handles=handles[1:], labels=labels[1:],loc='upper right', ncol=1,prop={'size': 6})
    #plt.legend(loc='upper right', bbox_to_anchor=(1.25, 0.5), ncol=3)
    plt.title(str('Per Capita Index - ' + i ))
    x = dt.datetime(2020, 5, 1)
    plt.annotate('Overperforming expectations', xy=(mdates.date2num(x), .01))
    plt.annotate('Underperforming expectations', xy=(mdates.date2num(x), .99))
    filename = 'charts/' + str(i) + '_6.png'
    fig = plt.savefig(filename,bbox_inches='tight')
    plt.close('all')
    chart_sheet.insert_image('J30',filename)
    print(i)




recent = df2['Date'].max().strftime('%x')




writer.save()

df3 = df2[df2['Date'] == today]
df3 = df3[['State','hospitalization_change']]
df3 = df3.sort_values('State')

df4 = df2[df2['Date'] == yesterday]
df4 = df4[['State','daily_deaths']]
df4 = df4.sort_values('daily_deaths',ascending=False).head()
df4 = df4.set_index('State')

df5 = df2[df2['Date'] == yesterday]
df5 = df5[['State','daily_hospitalized']]
df5 = df5.sort_values('daily_hospitalized',ascending=False).head()
df5 = df5.set_index('State')

biginc = df3[df3['hospitalization_change'] >= .25]
smallinc = df3[(df3['hospitalization_change'] < .25) & (df3['hospitalization_change'] >= 0)]
smalldec = df3[(df3['hospitalization_change'] < 0) & (df3['hospitalization_change'] >= -0.25)]
bigdec = df3[df3['hospitalization_change'] < -0.25]
notavail = df3[df3['hospitalization_change'].isnull()]

# EMAIL _______________________________________________________________________________________________
fromaddr = "ryangerda@gmail.com"
#toaddr = ["ryangerda@gmail.com","JGerda@ta-petro.com","deryda@roadrunner.com","lscarasso@gmail.com"]
toaddr = ["ryangerda@gmail.com"]

# instance of MIMEMultipart
msg = MIMEMultipart()
# storing the senders email address
msg['From'] = fromaddr
# storing the receivers email address
msg['To'] = ", ".join(toaddr)
# storing the subject
msg['Subject'] = "COVID Data Daily Update " + clock
# string to store the body of the mail
body = ("Attached is COVID data from Johns Hopkins University and was most recently published on " + recent +
        " \n \n WARNING: This data is only as good as the data states make available. There are instances when a week's worth of data is dumped at once, causing dramatic swings or inconsistencies." +
        " \n \n States where the 7-day rolling hospitalization has increased by more than 25%: " +
        " \n " + ' \n '.join(biginc['State'].tolist()) +
        " \n \n States where the 7-day rolling hospitalization has increased by 0-25%: " +
        " \n " + ' \n '.join(smallinc['State'].tolist()) +
        " \n \n States where the 7-day rolling hospitalization has decreased by 0-25%: " +
        " \n " + ' \n '.join(smalldec['State'].tolist()) +
        " \n \n States where the 7-day rolling hospitalization has decreased by more than 25%: " +
        " \n " + ' \n '.join(bigdec['State'].tolist()) +
        " \n \n States where the 7-day rolling hospitalization is NOT AVAILABLE: " +
        " \n " + ' \n '.join(notavail['State'].tolist()) +
        " \n \n States with highest death count yesterday: " +
        " \n " + '''\n{}'''.format(df4.to_string()) +
        " \n \n States with highest hospitalization count yesterday: " +
        " \n " + '''\n{}'''.format(df5.to_string()))
# attach the body with the msg instance
msg.attach(MIMEText(body, 'plain'))
# open the file to be sent
filename = 'COVID_' + str(clock) + '.xlsx'
attachment = open("/Users/ryangerda/PycharmProjects/corona/COVID_" + str(clock) + ".xlsx", "rb")
# instance of MIMEBase and named as p
p = MIMEBase('application', 'octet-stream')
# To change the payload into encoded form
p.set_payload((attachment).read())
# encode into base64
encoders.encode_base64(p)
p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
# attach the instance 'p' to instance 'msg'
msg.attach(p)
# creates SMTP session
s = smtplib.SMTP('smtp.gmail.com', 587)
# start TLS for security
s.starttls()
# Authentication
s.login(fromaddr, "rfvejwvbglccgepp")
# Converts the Multipart msg into a string
text = msg.as_string()
# sending the mail
s.sendmail(fromaddr, toaddr, text)
# terminating the session
s.quit()
print('Email Sent!')

print("--- %s minutes ---" % round((time.time() - start_time)/60,3))




