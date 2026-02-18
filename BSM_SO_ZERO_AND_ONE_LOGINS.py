# is code me ME REGION maine apna vala map nhii kiya hia usko map karenge to data sahi ho summary bhi sahi 
# usko map kar denge to summary bhi shai ho jayegi
# Mnapower ka path har week chaange karna hoga
import pandas as pd
import pandasql as ps
import os
from datetime import timedelta
path = r"C:\Users\ET0001301\Desktop\Manpower\ManpowerME & LAG - 16th Feb'26.xlsm"
df = pd.read_excel(path,sheet_name='Data')
query = " select * from df where Department = 'ME' and Role in ('BSM ME','SO ME')"
result = ps.sqldf(query,locals())
to_date = pd.Timestamp.now()
yes_date = to_date-timedelta(days=1)
yes_date = yes_date.strftime('%d-%b-%Y')
yes_date
save_pat = r"C:\Users\ET0001301\Pictures\Data"
file_name = f'Zero and 1 Logins as on {yes_date} .xlsx'
save_path = os.path.join(save_pat,file_name)
result = result.drop(columns= ['Official_Email','ReportingManager_Name','ReportingManager_Code','ReportingManagersManagerName','ReportingManagersManagerCode','SubDepartment_Code','Role1'])
result['Date_Of_Joining'] = pd.to_datetime(result['Date_Of_Joining'],errors = 'ignore')
result['Date_Of_Joining'].dtype
result.insert(result.columns.get_loc('Date_Of_Joining')+1,'Feb Joining Exlude',pd.NA)
from datetime import datetime
this_month = pd.Timestamp.now().strftime('%b')
month_number = pd.Timestamp.now().month
this_year = pd.Timestamp.now().year
result['Feb Joining Exlude'] = result['Date_Of_Joining'].apply(lambda x:this_month if ((x.month == month_number) and (x.year == this_year)) else "")
result['Date_Of_Joining'] = result['Date_Of_Joining'].apply(lambda x:x.date() if pd.notnull(x) else x)
mail_date = pd.Timestamp.now().date()
yest_date = mail_date-(timedelta(days=1))
def ordinal(n):
    if n%100>=10 and n%100<=13:
        suffix = 'th'
    else:
        suffix = {1:'st',2:'nd',3:'rd'}.get(n%10,'th')
    return str(n)+suffix
# today_date = f'{ordinal(mail_date.day)} {mail_date.strftime("%b %y")}'
yes_date = f'{ordinal(yest_date.day)} {yest_date.strftime("%b'%y")}' 
mis_path = fr"C:\Users\ET0001301\Desktop\ME-MIS_SUMMARY\ME-MIS_SUMMARY as on {yes_date}.xlsx"
mis = pd.read_excel(mis_path,sheet_name='Sheet1')
mis = mis[(mis['Logins'].notna())&(mis['Tranch']=='NO')]
mis['REGISTRATIONStage'] = mis['REGISTRATIONStage'].astype(str)
mis_pivot = pd.pivot_table(mis,
                           index = 'REGISTRATIONStage',
                           columns = 'MIS_STATUS',
                           values = 'IN_Cr',
                           aggfunc='count',
                           margins=True,
                           margins_name='Total',
                           fill_value=0
                           )
mis_pivot = mis_pivot.reset_index()
result = result[result['Feb Joining Exlude']==""]
result['Total Logins'] = result['Employee_Code'].map(mis_pivot.set_index('REGISTRATIONStage')['Total'])
result['Total Logins'] = result['Total Logins'].fillna(0)
result
import numpy as np
conditions = [result['Total Logins']==0,result['Total Logins']==1]
choices = [0,1]

result['zero and 1'] = np.select(conditions,choices,default=np.nan)
final_sum = pd.pivot_table(result,
                           index = 'ME Region',
                           columns = 'zero and 1',
                           values = 'Employee_Code',
                           aggfunc='count',
                           margins = True,
                           margins_name='Total',
                           fill_value=0
                           )
final_sum = final_sum.reset_index()
with pd.ExcelWriter(save_path,engine='openpyxl',mode = 'w') as writer:
    
    result.to_excel(writer,sheet_name = 'manpower',index=False)
    mis_pivot.to_excel(writer,sheet_name='pivot',index=False)
    final_sum.to_excel(writer,sheet_name='Summary',index = False)

