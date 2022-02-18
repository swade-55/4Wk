import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
from datetime import datetime, timedelta
from pandas.tseries.offsets import *
from openpyxl import load_workbook
import matplotlib.pyplot as plt





st.write("""
# Robesonia Daily Production Report Builder App
This app produces daily report for C&S Robesonia facility.
""")

st.sidebar.header('User Input Features')

st.sidebar.markdown("""
[Example Triceps file](https://github.com/swade-55/Robi/blob/main/production_triceps_labor_report.xlsx?raw=true)
""")

# Collects user input features into dataframe
triceps_week1 = st.sidebar.file_uploader("Upload 1st Triceps File", type=["xlsx"])
triceps_week2 = st.sidebar.file_uploader("Upload 2nd Triceps File", type=["xlsx"])
triceps_week3 = st.sidebar.file_uploader("Upload 3rd Triceps file", type=["xlsx"])
triceps_week4 = st.sidebar.file_uploader("Upload 4th Triceps file", type=["xlsx"])



st.sidebar.markdown("""
[Example Qlik file](https://github.com/swade-55/Robi/blob/main/Robi%20Hours.xlsx?raw=true)
""")
qlik_file = st.sidebar.file_uploader("Upload your input Qlik file", type=["xlsx"])

















check1 = st.sidebar.button("Analyze")


text_contents = '''
Foo, Bar
123, 456
789, 000
'''

if triceps_week1 is not None:
    df1 = pd.read_excel(triceps_week1)


if triceps_week2 is not None:
    df2 = pd.read_excel(triceps_week2)

if triceps_week3 is not None:
    df3 = pd.read_excel(triceps_week3)

if triceps_week4 is not None:
    df4 = pd.read_excel(triceps_week4)

if qlik_file is not None:
    df9 = pd.read_excel(qlik_file)
if check1:
    df9 = df9.drop(columns = ['Warehouse','Week Ending','Shift','Status','FT/PT','Units','Indirect Hours','Productivity','Performance','Engagements','GER'])
    df1 = df1.drop(df1.index[0])
    df1 = df1.drop(df1.index[0])
    df1.columns = df1.iloc[0]
    df1 = df1[1:]
    df2 = df2.drop(df2.index[0])
    df2 = df2.drop(df2.index[0])
    df2.columns = df2.iloc[0]
    df2 = df2[1:]
    df3 = df3.drop(df3.index[0])
    df3 = df3.drop(df3.index[0])
    df3.columns = df3.iloc[0]
    df3 = df3[1:]
    df4 = df4.drop(df4.index[0])
    df4 = df4.drop(df4.index[0])
    df4.columns = df4.iloc[0]
    df4 = df4[1:]
    df = pd.concat([df1,df2,df3,df4])
    def data(Triceps, Qlik):
        data = Triceps.copy()
        data['ACT_MINUTES'] = data['ACT_MINUTES'].astype(float, copy=False)
        data['STD_MINUTES'] = data['STD_MINUTES'].astype(str, copy=False)
        data['STD_MINUTES'] = data['STD_MINUTES'].astype(float, copy=False)
        data['EMPL_NUMBER'] = data['EMPL_NUMBER'].astype(int, copy=False)
        data['COMPLETED_CASES'] = data['COMPLETED_CASES'].astype(float, copy=False)
        data['IDLE_MIN'] = data['IDLE_MIN'].astype(float, copy=False)
        data['TASK'] = data['TASK'].astype(str, copy=False)
        #data['START_DATE_TIME'] = pd.to_datetime(data['START_DATE_TIME'])
        #start = datetime.strptime('10:00:00', '%H:%M:%S').time()
        #end = datetime.strptime('16:59:00', '%H:%M:%S').time()
        #data = data[data['START_DATE_TIME'].dt.time.between(start, end)]
        data['START_DATE_TIME'] = data['START_DATE_TIME'].astype(str, copy=False)
        data["Day"] = data['START_DATE_TIME'].str[0:10]
        numbers = data.groupby(['Day', 'EMPL_NUMBER', 'WHSE'], as_index=False).sum()
        numbers['Performance'] = numbers['STD_MINUTES'] / (numbers['ACT_MINUTES'] + numbers['IDLE_MIN']) * 100
        numbers['Day'] = pd.to_datetime(numbers['Day'])
        numbers['Day_of_Week'] = numbers['Day'].dt.day_name()
        Qlik['Date'] = pd.to_datetime(Qlik.Date)
        Qlik = Qlik.rename(columns = {'Total Hours':'Total_Hours'})
        numbers['WHSE'] = numbers['WHSE'].astype(int)
        numbers['WHSE'] = numbers['WHSE'].map({1: 'GDC', 2: 'PDC', 3: 'FDC'})
        data = Qlik.merge(numbers, how='inner', left_on=['Employee ID', 'Date', 'Commodity'], right_on=['EMPL_NUMBER', 'Day', 'WHSE'])
        data['Date'] = data['Date'].dt.strftime('%m/%d/%Y')
        data['Date'] = data['Date'].astype(str)
        data = data.rename(columns={'Commodity': 'Dept', 'Hire Date': 'DOH', 'Total Hours': 'Total_Hours'})
        data['DOH'] = data['DOH'].dt.strftime('%m/%d/%Y')
        data['DOH'] = data['DOH'].astype(str)
        data1 = data.groupby(['Position','Day_of_Week', 'Dept'], as_index=False).sum()
        data1 = data1.drop(columns=['Performance'])
        data1['Uptime'] = (data1['ACT_MINUTES']) / (data1['Total_Hours'] * 60) * 100
        data1['Performance'] = data1['STD_MINUTES'] / (data1['ACT_MINUTES'] + data1['IDLE_MIN']) * 100
        data1['Date'] = 'Indy'
        data1['Cases/Hour'] = data1['COMPLETED_CASES']/(data1['ACT_MINUTES']/60)
        data1['Day_of_Week'] = data1['Day_of_Week'].map({'Sunday':0,'Monday':1,'Tuesday':2,'Wednesday':3,'Thursday':4,'Friday':5,'Saturday':6})
        return data1

    def load(Triceps, Qlik):
        data = Triceps.copy()
        data['Pallets'] = 1
        data['ACT_MINUTES'] = data['ACT_MINUTES'].astype(float, copy=False)
        data['STD_MINUTES'] = data['STD_MINUTES'].astype(str, copy=False)
        data['STD_MINUTES'] = data['STD_MINUTES'].astype(float, copy=False)
        data['EMPL_NUMBER'] = data['EMPL_NUMBER'].astype(int, copy=False)
        data['COMPLETED_CASES'] = data['COMPLETED_CASES'].astype(float, copy=False)
        data['IDLE_MIN'] = data['IDLE_MIN'].astype(float, copy=False)
        data['TASK'] = data['TASK'].astype(str, copy=False)
        data['START_DATE_TIME'] = pd.to_datetime(data.START_DATE_TIME)
        data['START_DATE_TIME'] = pd.to_datetime(data.START_DATE_TIME) - timedelta(hours=5)
        data['START_DATE_TIME'] = data['START_DATE_TIME'].astype(str, copy=False)
        data["Day"] = data['START_DATE_TIME'].str[0:10]
        numbers = data.groupby(['Day', 'EMPL_NUMBER', 'WHSE'], as_index=False).sum()
        numbers['Performance'] = numbers['STD_MINUTES'] / (numbers['ACT_MINUTES'] + numbers['IDLE_MIN']) * 100
        numbers['Day'] = pd.to_datetime(numbers['Day'])
        numbers['Day_of_Week'] = numbers['Day'].dt.day_name()
        Qlik['Date'] = pd.to_datetime(Qlik.Date)
        numbers['WHSE'] = numbers['WHSE'].astype(int)
        numbers['WHSE'] = numbers['WHSE'].map({1: 'GDC', 2: 'PDC', 3: 'FDC'})
        data = Qlik.merge(numbers, how='inner', left_on=['Employee ID', 'Date'],right_on=['EMPL_NUMBER', 'Day'])
        data['Date'] = data['Date'].dt.strftime('%m/%d/%Y')
        data['Date'] = data['Date'].astype(str)
        data['Scans/Hour'] = (data['Pallets'] / data['Total Hours'])
        data = data.rename(columns={'Commodity': 'Dept', 'Hire Date': 'DOH', 'Total Hours': 'Total_Hours'})
        data['DOH'] = data['DOH'].dt.strftime('%m/%d/%Y')
        data['DOH'] = data['DOH'].astype(str)
        data1 = data.groupby(['Position', 'Dept','Day_of_Week'], as_index=False).sum()
        data1 = data1.drop(columns=['Scans/Hour'])
        data1['Uptime'] = (data1['ACT_MINUTES'] + data1['IDLE_MIN']) / (data1['Total_Hours'] * 60) * 100
        data1['Scans/Hour'] = (data1['Pallets'] / data1['Total_Hours'])
        data1['Date'] = 'Total'
        data1['Day_of_Week'] = data1['Day_of_Week'].map({'Sunday':0,'Monday':1,'Tuesday':2,'Wednesday':3,'Thursday':4,'Friday':5,'Saturday':6})
        mypiv = data1.pivot_table(index=['Position'],columns='Day_of_Week')[['Scans/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()
        return mypiv
    def fork(Triceps, Qlik):
        data = Triceps.copy()
        data['ACT_MINUTES']= data['ACT_MINUTES'].astype(float,copy=False)
        data['STD_MINUTES']= data['STD_MINUTES'].astype(str,copy=False)
        data['STD_MINUTES']= data['STD_MINUTES'].astype(float,copy=False)
        data['EMPL_NUMBER']= data['EMPL_NUMBER'].astype(int,copy=False)
        data['COMPLETED_CASES']= data['COMPLETED_CASES'].astype(float,copy=False)
        data['IDLE_MIN']= data['IDLE_MIN'].astype(float,copy=False)
        data['TASK']= data['TASK'].astype(str,copy=False)
        data['START_DATE_TIME']= data['START_DATE_TIME'].astype(str,copy=False)
        data["Day"] = data['START_DATE_TIME'].str[0:10]
        data['Pallets'] = 1
        numbers = data.groupby(['Day','EMPL_NUMBER'],as_index=False).sum()
        numbers['Day'] = pd.to_datetime(numbers.Day)
        numbers['Day'] =numbers['Day'].dt.strftime('%m/%d/%Y')
        numbers['Day'] = pd.to_datetime(numbers.Day)
        numbers['Day'] = pd.to_datetime(numbers['Day'])
        numbers['Day_of_Week'] = numbers['Day'].dt.day_name()
        Qlik['Date'] = pd.to_datetime(Qlik.Date)
        Qlik = Qlik.rename(columns = {'Total Hours':'Total_Hours'})
        data = Qlik.merge(numbers, how='inner', left_on=['Employee ID','Date'], right_on=['EMPL_NUMBER','Day'])
        data['Date'] =data['Date'].dt.strftime('%m/%d/%Y')
        data['Date'] = data['Date'].astype(str)
        data = data.rename(columns = {'Hire Date':'DOH'})
        data['DOH'] = pd.to_datetime(data.DOH)
        data['DOH'] =data['DOH'].dt.strftime('%m/%d/%Y')
        data['DOH'] = data['DOH'].astype(str)
        data = data.groupby(['Position','Day_of_Week','Commodity'],as_index=False).sum()
        data['Pallets/Hour'] = data['Pallets']/(data['ACT_MINUTES']/60)
        data['Uptime'] = (data['ACT_MINUTES']+data['IDLE_MIN'])/(data['Total_Hours']*60)*100
        data['Day_of_Week'] = data['Day_of_Week'].map({'Sunday':0,'Monday':1,'Tuesday':2,'Wednesday':3,'Thursday':4,'Friday':5,'Saturday':6})
        return data
    Puttriceps = df[df['JOB_CODE']=='PUT']
    Selecttriceps1 = df[df['JOB_CODE']=='CSL']
    Selecttriceps2 = df[df['JOB_CODE']=='CSE']
    Selecttriceps = Selecttriceps1.append(Selecttriceps2)
    Loadtriceps = df[df['JOB_CODE']=='LOD']
    Lettriceps = df[df['JOB_CODE']=='LET']
    forkhour3 = df9[df9['Position']=='Operator, Forklift - Step']
    forkhour1 = df9[df9['Position']=='Operator, Forklift']
    forkhour2 = df9[df9['Position']=='Forklift, Hourly, Freezer - Step']
    ForkQlik = forkhour1.append([forkhour3,forkhour2])
    selecthour3 = df9[df9['Position']=='Selector, Incentive, Freezer - Step']
    selecthour1 = df9[df9['Position']=='Selector, In Training']
    selecthour2 = df9[df9['Position']=='Selector, Incentive']       
    selecthour4 = df9[df9['Position']=='Selector, Incentive (ITT)']               
    SelectQlik = selecthour1.append([selecthour3,selecthour2,selecthour4])
    LoadQlik = df9[df9['Position']=='Loader - Step']
    Let = fork(Lettriceps, ForkQlik)
    sel = data(Selecttriceps, SelectQlik)
    loaders = load(Loadtriceps,LoadQlik)
    Let = Let.rename(columns = {'Pallets/Hour':'Letdowns/Hour'})
    GDClet = Let[Let['Commodity']=='GDC']
    PDClet = Let[Let['Commodity']=='PDC']
    FDClet = Let[Let['Commodity']=='FDC']

    GDClet = GDClet.pivot_table(index=['Position'],columns='Day_of_Week')[['Letdowns/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()
    PDClet = PDClet.pivot_table(index=['Position'],columns='Day_of_Week')[['Letdowns/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()
    FDClet = FDClet.pivot_table(index=['Position'],columns='Day_of_Week')[['Letdowns/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()
    GDCsel = sel[sel['Dept']=='GDC']
    PDCsel = sel[sel['Dept']=='PDC']
    FDCsel = sel[sel['Dept']=='FDC']

    GDCsel = GDCsel.pivot_table(index=['Position'],columns='Day_of_Week')[['Cases/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()
    PDCsel = PDCsel.pivot_table(index=['Position'],columns='Day_of_Week')[['Cases/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()
    FDCsel = FDCsel.pivot_table(index=['Position'],columns='Day_of_Week')[['Cases/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()
    Travel = df[df['JOB_CODE']=='TRV']
    Travel['START_DATE_TIME']= Travel['START_DATE_TIME'].astype(str,copy=False)
    Travel['Day'] = Travel['START_DATE_TIME'].str[0:10]
    Travel['Day'] = pd.to_datetime(Travel.Day)
    Travel['Day'] =Travel['Day'].dt.strftime('%m/%d/%Y')
    Travel['Day'] = pd.to_datetime(Travel.Day)
    Travel['Day'] = pd.to_datetime(Travel['Day'])
    Travel['Day_of_Week'] = Travel['Day'].dt.day_name()
    Travel = Travel.groupby(['Day_of_Week','WHSE'],as_index=False).sum()
    Travel = Travel.rename(columns = {'ACT_MINUTES':'TRV_MINUTES'})
    Travel = Travel.drop(columns = ['JOB_CODE','FACILITY','START_DATE_TIME','STD_MINUTES','IDLE_MIN','DELAY_MINUTES','COMPLETED_CUBE','COMPLETED_CASES','TASK','EMPL_NUMBER'])
    Travel['WHSE']= Travel['WHSE'].astype(int,copy=False)
    Travel['WHSE'] = Travel['WHSE'].map({1: 'GDC', 2: 'PDC', 3: 'FDC'})
    Travel['Day_of_Week'] = Travel['Day_of_Week'].map({'Sunday':0,'Monday':1,'Tuesday':2,'Wednesday':3,'Thursday':4,'Friday':5,'Saturday':6})
    Travel['Day_of_Week'] = Travel['Day_of_Week'].astype(int)
    Put = fork(Puttriceps,ForkQlik)
    Put = Put.merge(Travel, how='left', left_on=['Commodity','Day_of_Week'], right_on=['WHSE','Day_of_Week'])
    Put = Put.drop(columns = ['Pallets/Hour','Uptime'])
    Put['Total_MINUTES'] = Put['ACT_MINUTES']+Put['TRV_MINUTES']
    Put['Pallets/Hour'] = Put['Pallets']/(Put['Total_MINUTES']/60)
    Put['Class'] = 'Putaways/Hour'
    Put['Uptime'] = (Put['ACT_MINUTES']+Put['IDLE_MIN'])/(Put['Total_Hours']*60)*100
    Put = Put.rename(columns = {'Total Hours':'Total_Hours'})
    Put = Put.rename(columns = {'Pallets/Hour':'Putaways/Hour'})
    GDCPut = Put[Put['Commodity']=='GDC']
    PDCPut = Put[Put['Commodity']=='PDC']
    FDCPut = Put[Put['Commodity']=='FDC']

    GDCPut = GDCPut.pivot_table(index=['Position'],columns='Day_of_Week')[['Putaways/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()
    PDCPut = PDCPut.pivot_table(index=['Position'],columns='Day_of_Week')[['Putaways/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()
    FDCPut = FDCPut.pivot_table(index=['Position'],columns='Day_of_Week')[['Putaways/Hour','Uptime']].sort_values(by=['Position'], ascending=False).round()






    # Function to save all dataframes to one single excel
    def to_excel(df,df1,df2,df3,df4,df5,df6,df7,df8,df9):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=True, sheet_name='GDCsel')
        df1.to_excel(writer, index=True, sheet_name='PDCsel')
        df2.to_excel(writer, index=True, sheet_name='FDCsel')
        df3.to_excel(writer, index=True, sheet_name='GDClet')
        df4.to_excel(writer, index=True, sheet_name='PDClet')
        df5.to_excel(writer, index=True, sheet_name='FDClet')
        df6.to_excel(writer, index=True, sheet_name='GDCput')
        df7.to_excel(writer, index=True, sheet_name='PDCput')
        df8.to_excel(writer, index=True, sheet_name='FDCput')
        df9.to_excel(writer, index=True, sheet_name='loaders')

        workbook = writer.book
        #worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        #worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data
    df_xlsx = to_excel(GDCsel,PDCsel,FDCsel,GDClet,PDClet,FDClet,GDCPut,PDCPut,FDCPut,loaders)
    st.download_button(label='📥 Download Current Result', data=df_xlsx ,file_name= '4Wk.xlsx')

        

    st.subheader('GDC')
    st.write(GDCsel)

    st.subheader('PDC')
    st.write(PDCsel)

    st.subheader('FDC')
    st.write(FDCsel)



