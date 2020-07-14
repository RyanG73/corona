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
df['daily_newcases'] = df.groupby('State')['Confirmed'].diff().fillna(0)
df['rolling_newcases'] = df['daily_newcases'].rolling(window=7).mean()
df['daily_deaths'] = df.groupby('State')['Deaths'].diff().fillna(0)
df['rolling_deaths'] = df['daily_deaths'].rolling(window=7).mean()
df['daily_recovered'] = df.groupby('State')['Recovered'].diff().fillna(0)
df['rolling_recovered'] = df['daily_recovered'].rolling(window=7).mean()
df['daily_tested'] = df.groupby('State')['People_Tested'].diff().fillna(0)
df['rolling_tested'] = df['daily_tested'].rolling(window=7).mean()
df['daily_hospitalized'] = df.groupby('State')['People_Hospitalized'].diff().fillna(0)
df['rolling_hospitalized'] = df['daily_hospitalized'].rolling(window=7).mean()
df['hospitalization_change'] = df.groupby('State')['rolling_hospitalized'].apply(lambda x: x.div(x.iloc[0]).subtract(7).mul(100))
df = df.set_index('Date')

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
ax.plot(df.index,df["rolling_hospitalized"], color='Red', label='7-Day Rolling Average')
ax.bar(df.index,df["daily_hospitalized"], color='Blue', label='Daily Count')
plt.xticks(rotation=45)
ax.set_xlabel("Date")
ax.set_ylabel("Number Hospitalized")
ax.legend(loc='upper right')
plt.title(str('Ohio'))
plt.show()

writer = pd.ExcelWriter('data.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Data', index=False)
writer.save()