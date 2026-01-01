import pandas as pd
import os


path = r"C:\Users\Ravi Gupta\Downloads\ManpowerME & LAG.xlsm"
df = pd.read_excel(path,sheet_name = 'Data')

dp = df[(df['Department']=='ME') & (df['Role'].isin ({'ABM ME','CBM ME','TL','SO ME','BSM ME'}))]

dp = dp.drop(['Entity','Final Location','LAG Region','Date_Of_Joining','Department','Official_Email','ReportingManager_Name','Location','State_Name','Zone',
              'ReportingManagersManagerName','ReportingManagersManagerCode','SubDepartment_Code','District',
              'Role1','ME Region','ReportingManager_Code'],axis = 1,errors = 'coerce')
dp.insert(6,'ME Region',pd.NA)
dp.insert(8,'Role 1',pd.NA)

dp['Employee_Name'] = dp['Employee_Name'].str.title()
# print(dp.columns)

abm_cbm_path = r"C:\Users\Ravi Gupta\Downloads\ABM CBM Logins as on 29th Dec'25 .xlsx"
abm_cbm_file = pd.read_excel(abm_cbm_path,sheet_name = 'Data')


final_branch_unique = abm_cbm_file.drop_duplicates(subset = 'Final Branch',keep = 'first')

dp['ME Region'] = dp['Final Branch'].map(final_branch_unique.set_index('Final Branch')['ME Region'])

dp['Role 1'] = dp.apply(lambda x:x['Role'] if x['Role'] in {'BSM ME','SO ME'} else "",axis = 1)


# dp = dp.rename(columns = {'ReportingManager_Code':'RM CODE 1'})
# # dp.insert(len(dp.columns),'RM Name 1',pd.NA)
# # dp.insert(len(dp.columns),'RM Desig 1',pd.NA)


for i in range(1,4):
    dp.insert(len(dp.columns),f'RM CODE {i}',pd.NA)
    dp.insert(len(dp.columns),f'RM NAME {i}',pd.NA)
    dp.insert(len(dp.columns),f'RM DESIG {i}',pd.NA)

dp['RM CODE 1']=dp['Employee_Code'].map(df.set_index('Employee_Code')['ReportingManager_Code']) 
dp['RM NAME 1'] = dp['RM CODE 1'].map(df.set_index( 'Employee_Code')['Employee_Name'])
dp['RM DESIG 1']=dp['RM CODE 1'].map(df.set_index('Employee_Code')['Role'])

dp['RM CODE 2'] = dp['RM CODE 1'].map(df.set_index('Employee_Code')['ReportingManager_Code'])
dp['RM NAME 2'] = dp['RM CODE 2'].map(df.set_index('Employee_Code')['Employee_Name'])
dp['RM DESIG 2'] =dp['RM CODE 2'].map(df.set_index('Employee_Code')['Role'])

dp['RM CODE 3'] = dp['RM CODE 2'].map(df.set_index('Employee_Code')['ReportingManager_Code'])
dp['RM NAME 3'] = dp['RM CODE 3'].map(df.set_index('Employee_Code')['Employee_Name'])
dp['RM DESIG 3'] = dp['RM CODE 3'].map(df.set_index('Employee_Code')['Role'])    


for i in range(1,4):
    dp[f'RM CODE {i}'] = dp.apply(lambda x:'-' if x[f'RM DESIG {i}'] not in {'ABM ME','CBM ME','TL'} else x[f'RM CODE {i}'],axis = 1)
    dp[f'RM NAME {i}'] = dp.apply(lambda x:'-' if x[f'RM DESIG {i}'] not in {'ABM ME','CBM ME','TL'} else x[f'RM NAME {i}'],axis = 1)
    dp[f'RM DESIG {i}'] = dp.apply(lambda x:'-' if x[f'RM DESIG {i}'] not in {'ABM ME','CBM ME','TL'} else x[f'RM DESIG {i}'],axis = 1)
    
    
dp.insert(len(dp.columns),'TL CODE',pd.NA)
dp.insert(len(dp.columns),'TL NAME',pd.NA)
dp.insert(len(dp.columns),'ABM CODE',pd.NA)
dp.insert(len(dp.columns),'ABM NAME',pd.NA)
dp.insert(len(dp.columns),'CBM CODE',pd.NA)
dp.insert(len(dp.columns),'CBM NAME',pd.NA)
dp.insert(len(dp.columns),'DIRECT RH',pd.NA)


f1 = dp[(dp['RM DESIG 1']=='TL') & (dp['RM DESIG 2']=='ABM ME') & (dp['RM DESIG 3']=='CBM ME')]
f1_index = f1.index
dp.loc[f1_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','CBM CODE','CBM NAME']] = dp.loc[f1_index,['RM CODE 1','RM NAME 1','RM CODE 2','RM NAME 2','RM CODE 3','RM NAME 3']]
dp.loc[f1_index,'DIRECT RH'] = '-'


f2 = dp[(dp.get('TL NAME').isna()) & (dp.get('RM DESIG 1') == 'TL') & (dp.get('RM DESIG 2')=='ABM ME')]
f2_index = f2.index
dp.loc[f2_index,['TL CODE','TL NAME','ABM CODE','ABM NAME']] = dp.loc[f2_index,['RM CODE 1','RM NAME 1','RM CODE 2','RM NAME 2']]
dp.loc[f2_index,['CBM CODE','CBM NAME','DIRECT RH']]='-'

f3 = dp[(dp.get('RM DESIG 1') == 'TL') & (dp.get('TL CODE').isna()) & (dp.get('RM DESIG 2') == 'CBM ME')]
f3_index = f3.index
dp.loc[f3_index,['TL CODE','TL NAME','CBM CODE','CBM NAME']] = dp.loc[f3_index,['RM CODE 1','RM NAME 1','RM CODE 2','RM NAME 2']]
dp.loc[f3_index,['ABM CODE','ABM NAME','DIRECT RH']] = "-"

#  there can be have this condition also
f4  = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='TL') & (dp.get('RM DESIG 2')=='TL')]
f4_index = f4.index
dp.loc[f4_index,['TL CODE','TL NAME']] = dp.loc[f4_index,['RM CODE 1','RM NAME 1']]

#  there can be have this condition also
f5  = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='TL') & (dp.get('RM DESIG 3')=='TL')]
f5_index = f5.index
dp.loc[f5_index,['TL CODE','TL NAME']] = dp.loc[f5_index,['RM CODE 1','RM NAME 1']]



 

# f2 = dp[(dp.get('TL CODE').isna()) & ( dp.get('RM DESIG 1')=='TL')]
# # dp.loc[f1,'TL Name'] = dp.loc[f1,'RM Desig 1']









    
save_path = r"C:\Users\Ravi Gupta\Downloads\HL data"
os.makedirs(save_path,exist_ok = True)
file_name ='cleand_data.xlsx'
final_path = os.path.join(save_path,file_name)




dp.to_excel(final_path,index = False)
print("save successfull")
    



    
