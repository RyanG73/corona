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

#df['Last_Update'] = pd.to_datetime(df['Last_Update']).dt.date
df['Last_Update'] = pd.to_datetime(df['Last_Update']).dt.normalize()
df = df.rename(columns={'Province_State':'State','Last_Update':'Date'})
#df = df['State'].isin(['Ohio'])
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
ohio = ['Ohio']
df = df[df['State'].isin(ohio)]
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
df2 = df2.set_index('Date')
"""
sns.set_color_codes("pastel")
for i in states:
    fig = plt.figure()
    _ = ax1 = fig.add_subplot(111)
    _ = sns.lineplot(x='Date', y='rolling_deaths', data=df)
    _ = ax2 = ax1.twinx()
    _ = sns.barplot(x='Date', y='daily_deaths', data=df, color="b")
    #ax2.grid(False)
    plt.show()
"""

fig, ax = plt.subplots()
ax.plot(df2.index,df2["rolling_hospitalized"], color='Red', label='7-Day Rolling Average')
ax.bar(df2.index,df2["daily_hospitalized"], color='Blue', label='Daily Count')
plt.xticks(rotation=45)
ax.set_xlabel("Date")
ax.set_ylabel("Number Hospitalized")
ax.legend(loc='upper right')
plt.title(str(df['State'].unique()))


axe = sns.lineplot(x=df2.index, y="hospitalization_change", data=df2)
axe.axhline(0, ls='--',color='black')

plt.show()

writer = pd.ExcelWriter('data.xlsx', engine='xlsxwriter')
df2.to_excel(writer, sheet_name='Data', index=True)
writer.save()

# TODO: iron out charts
# TODO: set up emails