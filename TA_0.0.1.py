import git
import pandas as pd
import numpy as np
import glob
import datetime as dt
import seaborn as sns
import matplotlib.pyplot as plt

repo = git.Repo('/Users/ryangerda/PycharmProjects/corona/COVID-19')
repo.remotes.origin.pull()

path = r'/Users/ryangerda/PycharmProjects/corona/COVID-19/csse_covid_19_data/csse_covid_19_daily_reports_us' # use your path
all_files = glob.glob(path + "/*.csv")
li = []

df = pd.concat((pd.read_csv(f) for f in all_files))

df['Last_Update'] = pd.to_datetime(df['Last_Update']).dt.normalize()
df = df.rename(columns={'Province_State':'State','Last_Update':'Date'})
states = ['Alabama','Arizona', 'Arkansas',
           'California', 'Colorado', 'Connecticut', 'Delaware',
           'Florida', 'Georgia',
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
states = ['Utah','Ohio']
df = df[df['State'].isin(states)]
df = df.sort_values('Date')
df = df.reset_index(drop=True)
df2 = pd.DataFrame()
for i in states:
    d = df.loc[df['State'] == i].copy()
    d['daily_newcases'] = d['Confirmed'].diff().fillna(0)
    d['rolling_newcases'] = d['daily_newcases'].rolling(window=7).mean()
    d['daily_deaths'] = d['Deaths'].diff().fillna(0)
    d['rolling_deaths'] = d['daily_deaths'].rolling(window=7).mean()
    d['daily_recovered'] = d['Recovered'].diff().fillna(0)
    d['rolling_recovered'] = d['daily_recovered'].rolling(window=7).mean()
    d['daily_tested'] = d['People_Tested'].diff().fillna(0)
    d['rolling_tested'] = d['daily_tested'].rolling(window=7).mean()
    d['daily_hospitalized'] = d['People_Hospitalized'].diff().fillna(0)
    d['rolling_hospitalized'] = d['daily_hospitalized'].rolling(window=7).mean()
    d['hospitalization_change'] = d['rolling_hospitalized'] / d['rolling_hospitalized'].shift(7) - 1
    df2 = df2.append(d)
    #df2 = df2.set_index('Date')

df2 = df2[(df2['Date'] > '2020-04-20')]

writer = pd.ExcelWriter('data.xlsx', engine='xlsxwriter')
df2.to_excel(writer, sheet_name='Data', index=True)
workbook = writer.book

for e,s in enumerate(states):
    print(s)
    plt.close()
    d = df2.loc[df2['State'] == s].copy()
    print(d.State.value_counts())
    empty = pd.DataFrame()
    empty.to_excel(writer,sheet_name=str(s))
    chart_sheet = writer.sheets[str(s)]
    # CHART 1
    fig = plt.figure(figsize=(5,5))
    #fig, ax = plt.subplots(figsize=(5,5))
    plt.plot(df2['Date'],df2["rolling_tested"], color='Red', label='7-Day Rolling Average')
    plt.bar(df2['Date'],df2["daily_tested"], color='Blue', label='Daily Count')
    #ax.plot(df2['Date'],df2["rolling_tested"], color='Red', label='7-Day Rolling Average')
    #ax.bar(df2['Date'],df2["daily_tested"], color='Blue', label='Daily Count')
    #plt.xticks(rotation=20)
    #ax.set_xlabel("Date")
    #ax.set_ylabel("Number Tested")
    plt.xticks(rotation=20)
    plt.xlabel("Date")
    plt.ylabel("Number Tested")
    plt.legend(loc='upper right')
    plt.title(str('Testing'))
    plt.show()
    filename = 'charts/' + str(s) + str(e) + '.png'
    print(filename)
    plt.savefig(filename)
    #chart_sheet.insert_image('A1',filename)
    plt.close('all')
    print(s)
    """
    # CHART 2
    fig, ax = plt.subplots(figsize=(5,5))
    ax.plot(df2['Date'],df2["rolling_hospitalized"], color='Red', label='7-Day Rolling Average')
    ax.bar(df2['Date'],df2["daily_hospitalized"], color='Blue', label='Daily Count')
    plt.xticks(rotation=20)
    ax.set_xlabel("Date")
    ax.set_ylabel("Number Hospitalized")
    ax.legend(loc='upper right')
    plt.title(str('Hospitalizations'))
    plt.savefig(r'/Users/ryangerda/PycharmProjects/corona/charts/' + str(i) + 'figure2.png')
    chart_sheet.insert_image('I1',r'/Users/ryangerda/PycharmProjects/corona/charts/' + str(i) + 'figure2.png')
    print(i)
    # CHART 3
    fig, ax = plt.subplots(figsize=(5,5))
    ax.plot(df2['Date'], df2["hospitalization_change"], color='Red')
    #axe = sns.lineplot(x=df2['Date'], y="hospitalization_change", data=df2)
    ax.axhline(0, ls='--',color='black')
    plt.xticks(rotation=20)
    ax.set_xlabel("Date")
    ax.set_ylabel("Percent Change")
    plt.title(str('7-Day Rolling Change in Hospitalization'))
    plt.savefig(r'/Users/ryangerda/PycharmProjects/corona/charts/' + str(i) + 'figure3.png')
    chart_sheet.insert_image('I26',r'/Users/ryangerda/PycharmProjects/corona/charts/' + str(i) + 'figure3.png')
    print(i)
    # CHART 4
    fig, ax = plt.subplots(figsize=(5,5))
    ax.plot(df2['Date'],df2["rolling_newcases"], color='Red', label='7-Day Rolling Average')
    ax.bar(df2['Date'],df2["daily_newcases"], color='Blue', label='Daily Count')
    plt.xticks(rotation=20)
    ax.set_xlabel("Date")
    ax.set_ylabel("Count")
    ax.legend(loc='upper right')
    plt.title(str('New Cases'))
    plt.savefig(r'/Users/ryangerda/PycharmProjects/corona/charts/_' + str(i) + 'figure4.png')
    chart_sheet.insert_image('A26',r'/Users/ryangerda/PycharmProjects/corona/charts/_' + str(i) + 'figure4.png')
    print(i)
    """
plt.figure().clear()
plt.close(fig)
plt.close('all')
writer.save()

# TODO: iron out charts
# TODO: set up emails