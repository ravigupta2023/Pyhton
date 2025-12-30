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

# dp['ME Region'] = dp['Final Branch'].map(abm_cbm_file.set_index('Final Branch')['ME Region'])

# abm = abm_cbm_file['Final Branch','ME Region']

# dp = dp.merge(abm_cbm_file[['Final Branch','ME Region']], on = 'Final Branch',how = 'left')

final_branch_unique = abm_cbm_file.drop_duplicates(subset = 'Final Branch',keep = 'first')

dp['ME Region'] = dp['Final Branch'].map(final_branch_unique.set_index('Final Branch')['ME Region'])

dp['Role 1'] = dp.apply(lambda x:x['Role'] if x['Role'] in {'BSM ME','SO ME'} else "",axis = 1)


# dp = dp.rename(columns = {'ReportingManager_Code':'RM CODE 1'})
# # dp.insert(len(dp.columns),'RM Name 1',pd.NA)
# # dp.insert(len(dp.columns),'RM Desig 1',pd.NA)


for i in range(1,4):
    dp.insert(len(dp.columns),f'RM Code {i}',pd.NA)
    dp.insert(len(dp.columns),f'RM Name {i}',pd.NA)
    dp.insert(len(dp.columns),f'RM Desig {i}',pd.NA)

dp['RM Code 1']=dp['Employee_Code'].map(df.set_index('Employee_Code')['ReportingManager_Code']) 
dp['RM Name 1'] = dp['RM Code 1'].map(df.set_index( 'Employee_Code')['Employee_Name'])
dp['RM Desig 1']=dp['RM Code 1'].map(df.set_index('Employee_Code')['Role'])

dp['RM Code 2'] = dp['RM Code 1'].map(df.set_index('Employee_Code')['ReportingManager_Code'])
dp['RM Name 2'] = dp['RM Code 2'].map(df.set_index('Employee_Code')['Employee_Name'])
dp['RM Desig 2'] =dp['RM Code 2'].map(df.set_index('Employee_Code')['Role'])

dp['RM Code 3'] = dp['RM Code 2'].map(df.set_index('Employee_Code')['ReportingManager_Code'])
dp['RM Name 3'] = dp['RM Code 3'].map(df.set_index('Employee_Code')['Employee_Name'])
dp['RM Desig 3'] = dp['RM Code 3'].map(df.set_index('Employee_Code')['Role'])    


for i in range(1,4):
    dp[f'RM Code {i}'] = dp.apply(lambda x:'-' if x[f'RM Desig {i}'] not in {'ABM ME','CBM ME','TL'} else x[f'RM Code {i}'],axis = 1)
    dp[f'RM Name {i}'] = dp.apply(lambda x:'-' if x[f'RM Desig {i}'] not in {'ABM ME','CBM ME','TL'} else x[f'RM Name {i}'],axis = 1)
    dp[f'RM Desig {i}'] = dp.apply(lambda x:'-' if x[f'RM Desig {i}'] not in {'ABM ME','CBM ME','TL'} else x[f'RM Desig {i}'],axis = 1)
    
    
dp.insert(len(dp.columns),'TL Name',pd.NA)
dp.insert(len(dp.columns),'TL Code',pd.NA)
dp.insert(len(dp.columns),'ABM Code',pd.NA)
dp.insert(len(dp.columns),'ABM Name',pd.NA)
dp.insert(len(dp.columns),'CBM Code',pd.NA)
dp.insert(len(dp.columns),'CBM Name',pd.NA)
dp.insert(len(dp.columns),'Direct RH',pd.NA)


f1 = dp[(dp['RM Desig 1']=='TL') & (dp['RM Desig 2']=='ABM ME') & (dp['RM Desig 3']=='CBM ME')]
f1_index = f1.index
dp.loc[f1_index,'TL Name'] = f1['RM Name 1']
dp.loc[f1_index,'TL Code'] = f1['RM Code 1']
dp.loc[f1_index,'ABM Name'] = f1['RM Name2']
# dp.loc[f1,'TL Name'] = dp.loc[f1,'RM Desig 1']









    
save_path = r"C:\Users\Ravi Gupta\Downloads\HL data"
os.makedirs(save_path,exist_ok = True)
file_name ='cleand_data.xlsx'
final_path = os.path.join(save_path,file_name)




dp.to_excel(final_path,index = False)
    



    
