import pandas as pd
import os


path = r"C:\Users\et0001301\Pictures\ManpowerME & LAG-29th Dec'25.xlsm"
df = pd.read_excel(path,sheet_name = 'Data')

dp = df[(df['Department']=='ME') & (df['Role'].isin ({'ABM ME','CBM ME','TL','SO ME','BSM ME'}))]

dp = dp.drop(['Entity','Final Location','LAG Region','Date_Of_Joining','Department','Official_Email','ReportingManager_Name',
              'ReportingManagersManagerName','ReportingManagersManagerCode','SubDepartment_Code',
              'Role1','ME Region'],axis = 1)
dp.insert(6,'ME Region',pd.NA)
dp.insert(8,'Role 1',pd.NA)




dp = dp.rename(columns = {'ReportingManager_Code':'RM CODE 1'})
dp.insert(len(dp.columns),'RM Name 1',pd.NA)
dp.insert(len(dp.columns,'RM Desig 1'),pd.NA)


for i in range(1,4):
    dp.insert(len(dp.columns),f'RM Name {i}',pd.NA)
    dp.insert(len(dp.columns),f'RM Desig {i}',pd.NA)
    dp.insert(len(dp.columns),f'RM Code{i}',pd.NA)



    
