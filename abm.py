
import pandas as pd
import os


path = r"C:\Users\ET0001301\Pictures\ManpowerME & LAG-29th Dec'25.xlsm"
df = pd.read_excel(path,sheet_name = 'Data')

dp = df[(df['Department']=='ME') & (df['Role'].isin ({'ABM ME','CBM ME','TL','SO ME','BSM ME'}))]

dp = dp.drop(['Entity','Final Location','LAG Region','Date_Of_Joining','Department','Official_Email','ReportingManager_NAME','Location','State_NAME','Zone',
              'ReportingManagersManagerNAME','ReportingManagersManagerCODE','SubDepartment_CODE','District',
              'Role1','ME Region','ReportingManager_CODE'],axis = 1,errors = 'ignore')
dp.insert(6,'ME Region',pd.NA)
dp.insert(8,'Role 1',pd.NA)

dp['Employee_Name'] = dp['Employee_Name'].str.title()
# print(dp.columns)

abm_cbm_path = r"C:\Users\ET0001301\Pictures\ABM CBM Logins as on 30th Dec'25 .xlsx"
abm_cbm_file = pd.read_excel(abm_cbm_path,sheet_name = 'Data')

# dp['ME Region'] = dp['Final Branch'].map(abm_cbm_file.set_index('Final Branch')['ME Region'])

# abm = abm_cbm_file['Final Branch','ME Region']

# dp = dp.merge(abm_cbm_file[['Final Branch','ME Region']], on = 'Final Branch',how = 'left')

final_branch_unique = abm_cbm_file.drop_duplicates(subset = 'Final Branch',keep = 'first')

dp['ME Region'] = dp['Final Branch'].map(final_branch_unique.set_index('Final Branch')['ME Region'])

dp['Role 1'] = dp.apply(lambda x:x['Role'] if x['Role'] in {'BSM ME','SO ME'} else "",axis = 1)


# dp = dp.reNAME(columns = {'ReportingManager_CODE':'RM CODE 1'})
# # dp.insert(len(dp.columns),'RM NAME 1',pd.NA)
# # dp.insert(len(dp.columns),'RM DESIG 1',pd.NA)


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

# FILTER ON ALL ABM CBM AND TL
f1 = dp[(dp['RM DESIG 1']=='TL') & (dp['RM DESIG 2']=='ABM ME') & (dp['RM DESIG 3']=='CBM ME')]
f1_index = f1.index
dp.loc[f1_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','CBM CODE','CBM NAME']] = dp.loc[f1_index,['RM CODE 1','RM NAME 1','RM CODE 2','RM NAME 2','RM CODE 3','RM NAME 3']].values
dp.loc[f1_index,'DIRECT RH'] = '-'

#  FIlTER FOR 1 = TL AND 2 =ABM
f2 = dp[(dp.get('RM DESIG 1')=='TL') & (dp.get('RM DESIG 2') =='ABM ME') & (dp.get('TL CODE').isna())]
f2_index = f2.index
dp.loc[f2_index,['TL CODE','TL NAME','ABM CODE','ABM NAME']]=dp.loc[f2_index,['RM CODE 1','RM NAME 1','RM CODE 2','RM NAME 2']].values
dp.loc[f2_index,['CBM CODE','CBM NAME','DIRECT RH']] = '-'

# FILTER FOR 1 = TL AND 2 = CBM
f3 = dp[(dp.get('RM DESIG 1') == 'TL') & (dp.get('RM DESIG 2')=='CBM ME') & (dp.get('TL CODE').isna())]
f3_index = f3.index
dp.loc[f3_index,['TL CODE','TL NAME','CBM CODE','CBM NAME']] = dp.loc[f3_index,['RM CODE 1','RM NAME 1','RM CODE 2','RM NAME 2']].values
dp.loc[f3_index,['ABM CODE','ABM NAME','DIRECT RH']] = '-'

# 1 = TL AND 2 = TL
f4 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1') =='TL') & (dp.get('RM DESIG 2')=='TL')]
f4_index  = f4.index
dp.loc[f4_index,['TL CODE','TL NAME']] = dp.loc[f4_index,['RM CODE 1','RM NAME 1']].values
dp.loc[f4_index,['ABM CODE','ABM NAME','CBM CODE','CBM NAME','DIRECT RH']]  ='-'

# 1 = TL AND 3 = TL
f5 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1') =='TL') & (dp.get('RM DESIG 3')=='TL')]
f5_index = f5.index
dp.loc[f5_index,['TL CODE','TL NAME']] = dp.loc[f5_index,['RM CODE 1','RM NAME 1']].values
dp.loc[f4_index,['ABM CODE','ABM NAME','CBM CODE','CBM NAME','DIRECT RH']]  ='-'

# 1 = TL AND 3 = ABM
f6 = dp[(dp.get('RM DESIG 1')=='TL') & (dp.get('RM DESIG 3')=='ABM ME') & (dp.get('TL CODE').isna())]
f6_index= f6.index
dp.loc[f6_index,['TL CODE','TL NAME','ABM CODE','ABM NAME']] = dp.loc[f6_index,['RM CODE 1','RM NAME 1','RM CODE 3','RM NAME 3']].values
dp.loc[f6_index,['CBM CODE','CBM NAME','DIRECT RH']] ='-'

# # 1 = TL AND 3 = CBM
f7 = dp[(dp.get('RM DESIG 1')=='TL') & (dp.get('RM DESIG 3')=='CBM ME') & (dp.get('TL CODE').isna())]
f7_index = f7.index
dp.loc[f7_index,['TL CODE','TL NAME','CBM CODE','CBM NAME']] = dp.loc[f7_index,['RM CODE 1','RM NAME 1','RM CODE 3','RM NAME 3']].values
dp.loc[f7_index,['ABM CODE','ABM NAME','DIRECT RH']]='-'

f8 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='TL') & (dp.get('RM DESIG 2')=='-') & (dp.get('RM DESIG 2') =='-')]
f8_index = f8.index
dp.loc[f8_index,['TL CODE','TL NAME']] = dp.loc[f8_index,['RM CODE 1','RM NAME 1']].values
dp.loc[f8_index,['ABM CODE','ABM NAME','CBM CODE','CBM NAME','DIRECT RH']] = '-'

f9 = dp[(dp.get('TL CODE').isna())& (dp.get('RM DESIG 1')=='ABM ME') & (dp.get('RM DESIG 2')=='CBM ME')]
f9_index = f9.index
dp.loc[f9_index,['ABM CODE','ABM NAME','CBM CODE','CBM NAME']] = dp.loc[f9_index,['RM CODE 1','RM NAME 1','RM CODE 2','RM NAME 2']].values
dp.loc[f9_index,['TL CODE','TL NAME','DIRECT RH']]='-'

f10 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='ABM ME') & (dp.get('RM DESIG 3')=='CBM ME')] 
f10_index = f10.index
dp.loc[f10_index,['ABM CODE','ABM NAME','CBM CODE','CBM NAME']] = dp.loc[f10_index,['RM CODE 1','RM NAME 1','RM CODE 3','RM NAME 3']].values
dp.loc[f10_index,['TL CODE','TL NAME','DIRECT RH']] ='-'

f11 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='ABM ME') & (dp.get('RM DESIG 2')=='ABM ME')]
f11_index = f11.index
dp.loc[f11_index,['ABM CODE','ABM NAME']] = dp.loc[f11_index,['RM CODE 1','RM NAME 1']].values
dp.loc[f11_index,['TL CODE','TL NAME','CBM CODE','CBM NAME','DIRECT RH']]  = '-'

f12 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='ABM ME') & dp.get('RM DESIG 3')=='ABM ME']
f12_index = f12.index
dp.loc[f12_index,['ABM CODE','ABM NAME']] = dp.loc[f12_index,['RM CODE 1','RM NAME 1']].values
dp.loc[f12_index,['TL CODE','TL NAME','CBM CODE','CBM NAME','DIRECT RH']] = '-'

f13 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1') =='ABM ME') & (dp.get('RM DESIG 2')=='-') & (dp.get('RM DESIG 3')=='-')]
f13_index = f13.index
dp.loc[f13_index,['ABM CODE','ABM NAME']] = dp.loc[f13_index,['RM CODE 1','RM NAME 1']].values
dp.loc[f13_index,['TL CODE','TL NAME','CBM CODE','CBM NAME','DIRECT RH']] = '-'

f14 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='CBM ME') & (dp.get('RM DESIG 2')=='CBM ME')]
f14_index = f14.index
dp.loc[f14_index,['CBM CODE','CBM NAME']]= dp.loc[f14_index,['RM CODE 1','RM NAME 1']].values
dp.loc[f14_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','DIRECT RH']] = '-'

f15 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='CBM ME') & (dp.get('RM DESIG 3')=='CBM ME')]
f15_index = f15.index
dp.loc[f15_index,['CBM CODE','CBM NAME']] = dp.loc[f15_index,['RM CODE 1','RM NAME 1']].values
dp.loc[f15_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','DIRECT RH']] = '-'

f16 = dp[dp.get('TL CODE').isna() & (dp.get('RM DESIG 1')=='CBM ME') & (dp.get('RM DESIG 2') =='-') & (dp.get('RM DESIG 3')=='-')]
f16_index = f16.index
dp.loc[f16_index,['CBM CODE','CBM NAME']] = dp.loc[f16_index,['RM CODE 1','RM NAME 1']].values
dp.loc[f16_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','DIRECT RH']] = '-'


# from here the second is starting

f17 = dp[(dp.get('RM DESIG 2')=='TL') & (dp.get('RM DESIG 3')=='ABM ME') & (dp.get('TL CODE').isna())]
f17_index = f17.index
dp.loc[f17_index,['TL CODE','TL NAME','ABM CODE','ABM NAME']] = dp.loc[f17_index,['RM CODE 2','RM NAME 2','RM CODE 3','RM NAME 3']].values
dp.loc[f17_index,['CBM CODE','CBM NAME','DIRECT RH']] = '-'

f22 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 2') == 'TL') & (dp.get('RM DESIG 3')=='CBM ME')]
f22_index = f22.index
dp.loc[f22_index,['TL CODE','TL NAME','CBM CODE','CBM NAME']] = dp.loc[f22_index,['RM CODE 2','RM NAME 2','RM CODE 3','RM NAME 3']].values
dp.loc[f22_index,['ABM CODE','ABM NAME','DIRECT RH']] = '-'

f18 = dp[(dp.get('RM DESIG 2')=='TL') & (dp.get('TL CODE').isna()) & (dp.get('RM DESIG 3')=='TL')]
f18_index = f18.index
dp.loc[f18_index,['TL CODE','TL NAME']] = dp.loc[f18_index,['RM CODE 2','RM NAME 2']].values
dp.loc[f18_index,['ABM CODE','ABM NAME','CBM CODE','CBM NAME','DIRECT RH']] = '-'

f19 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 2')=='TL') & (dp.get('RM DESIG 3') == '-')]
f19_index = f19.index
dp.loc[f19_index,['TL CODE','TL NAME']] = dp.loc[f19_index,['RM CODE 2','RM NAME 2']].values
dp.loc[f19_index,['ABM CODE','ABM NAME','CBM CODE','CBM NAME','DIRECT RH']] = '-'

#  from here the new variable i just defined

e1 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 2') == 'ABM ME') & (dp.get('RM DESIG 3')=='CBM ME')]
e1_index = e1.index
dp.loc[e1_index,['ABM CODE','ABM NAME','CBM CODE','CBM NAME']] = dp.loc[e1_index,['RM CODE 2','RM NAME 2','RM CODE 3','RM NAME 3']].values
dp.loc[e1_index,['TL CODE','TL NAME','DIRECT RH']] = '-'

e2 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 2')== 'ABM ME') & (dp.get('RM DESIG 3') == 'ABM ME')]
e2_index = e2.index
dp.loc[e2_index,['ABM CODE','ABM NAME']] = dp.loc[e2_index,['RM CODE 2','RM NAME 2']].values
dp.loc[e2_index,['TL CODE','TL NAME','CBM CODE','CBM NAME','DIRECT RH']] = '-'

e3 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 2')=='ABM ME') &  (dp.get('RM DESIG 3')=='-')]
e3_index = e3.index
dp.loc[e3_index,['ABM CODE','ABM NAME']]  = dp.loc[e3_index,['RM CODE 2','RM NAME 2']].values
dp.loc[e3_index,['TL CODE','TL NAME','CBM CODE','CBM NAME','DIRECT RH']] = '-'

e4 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 2')=='CBM ME') & (dp.get('RM DESIG 3') == '-')]
e4_index = e4.index
dp.loc[e4_index,['CBM CODE','CBM NAME']] = dp.loc[e4_index,['RM CODE 2','RM NAME 2']].values
dp.loc[e4_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','DIRECT RH']] = '-'

e5 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 2')=='CBM ME') & (dp.get('RM DESIG 3')=='CBM ME')]
e5_index = e5.index
dp.loc[e5_index,['CBM CODE',"CBM NAME"]] = dp.loc[e5_index,['RM CODE 2','RM NAME 2']].values
dp.loc[e5_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','DIRECT RH']] = '-'

#  last reporting manager now
c1 = dp[(dp.get('RM DESIG 3')=='ABM ME') & (dp.get('TL CODE').isna())]
c1_index = c1.index
dp.loc[c1_index,['ABM CODE','ABM NAME']] = dp.loc[c1_index,['RM CODE 3','RM NAME 3']].values
dp.loc[c1_index,['TL CODE','TL NAME','CBM CODE','CBM NAME','DIRECT RH']]='-'

c2 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 3')=='CBM ME')]
c2_index = c2.index
dp.loc[c2_index,['CBM CODE','CBM NAME']] = dp.loc[c2_index,['RM CODE 3','RM NAME 3']].values
dp.loc[c2_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','DIRECT RH']] = '-'



#  Now filling based on the role 

r1 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='-') & (dp.get('RM DESIG 2')=='-') & (dp.get('RM DESIG 3')=='-') & (dp.get('Role').isin({'BSM ME','SO ME'}))]
r1_index = r1.index
dp.loc[r1_index,'DIRECT RH'] = dp.loc[r1_index,'ME Region'].values
dp.loc[r1_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','CBM CODE','CBM NAME']] ='-'


r2 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='-') & (dp.get('RM DESIG 2')=='-') & (dp.get('RM DESIG 3')=='-') & (dp.get('Role')=='TL')]
r2_index = r2.index
dp.loc[r2_index,['TL CODE','TL NAME']] = dp.loc[r2_index,['Employee_Code','Employee_Name']].values
dp.loc[r2_index,['ABM CODE','ABM NAME','CBM CODE','CBM NAME','DIRECT RH']]='-'


r3 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='-') & (dp.get('RM DESIG 2')=='-') & (dp.get('RM DESIG 3')=='-') & (dp.get('Role')=='ABM ME')]
r3_index = r3.index
dp.loc[r3_index,['ABM CODE','ABM NAME']] = dp.loc[r3_index,['Employee_Code','Employee_Name']].values
dp.loc[r3_index,['TL CODE','TL NAME','CBM CODE','CBM NAME','DIRECT RH']] = '-'


r4 = dp[(dp.get('TL CODE').isna()) & (dp.get('RM DESIG 1')=='-') & (dp.get('RM DESIG 2')=='-') & (dp.get('RM DESIG 3')=='-') & (dp.get('Role')=='CBM ME')]
r4_index = r4.index
dp.loc[r4_index,['CBM CODE','CBM NAME']] = dp.loc[r4_index,['Employee_Code','Employee_Name']].values
dp.loc[r4_index,['TL CODE','TL NAME','ABM CODE','ABM NAME','DIRECT RH']] = '-'

# NOW THE TL IS MAPPING INTO THE ABM

t1 = dp[(dp.get('TL CODE') != "-") & (dp.get('ABM CODE')=='-')]
t1_index = t1.index
dp.loc[t1_index,['ABM CODE','ABM NAME']]  = dp.loc[t1_index,['TL CODE','TL NAME']]
dp.loc[t1_index,['TL CODE','TL NAME']] = '-'

# CAPITALIZING THE NAMES
dp['ABM NAME'] = dp['ABM NAME'].str.title()
dp['TL NAME'] = dp['TL NAME'].str.title()
dp['CBM NAME'] = dp['CBM NAME'].str.title()
dp['RM NAME 1'] = dp['RM NAME 1'].str.title()
dp['RM NAME 2'] = dp['RM NAME 2'].str.title()
dp['RM NAME 3'] = dp['RM NAME 3'].str.title()


# Adding the Value columns
col_value = ['Approved','Declined','Disbursed','OnHold','WIP','Total Logins','With Logins','Zero Logins','App count',
             'App Amt','Disb Count','Disb Amt','Total App Count','Total App Amt']
start_index = 2
for index,col_name in enumerate(col_value):
    dp.insert(start_index+index,col_name,pd.NA)



pivot = pd.pivot_table(dp,
            #    index = 'Employee_Code',
                values = ['Role 1','Total Logins','With Logins','Zero Logins','Total App Count','Total App Amt','Disb Count','Disb Amt'],
                index = ['Final Region','ME Region','CBM CODE','CBM NAME','ABM CODE','ABM NAME','TL CODE','TL NAME'],
                aggfunc = ['count','sum','count','count','sum','sum','sum','sum'],
                margins = True,
                fill_value = '-')




save_path = r"C:\Users\ET0001301\Pictures\ABM"
os.makedirs(save_path,exist_ok = True)
file_NAME ='cleand_data.xlsx'
final_path = os.path.join(save_path,file_NAME)


with pd.ExcelWriter(final_path,engine = 'openpyxl',mode = 'w') as writer:
    dp.to_excel(writer,sheet_name = 'data',index = False)
    pivot.to_excel(writer,sheet_name = 'summary')
print('save Successfull')    
 


    
