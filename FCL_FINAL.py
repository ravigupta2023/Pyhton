import pandas as pd

# Just have to chenagt the path of the fcl there is just have to save the path of the old one

path = r"C:\Users\ET0001301\Desktop\Python\FCL WHOLE DATA\Open Ticket_Report_14-02-2025 _After Tagging.xlsx"

# Give the ticket path wherer that is
df = pd.read_excel(path,sheet_name='Sheet0')
df['Created Date'] = pd.to_datetime(df['Created Date'],dayfirst=True,errors='ignore')
df = df[df['Folder Level 2']=='FCL required']

twenty_date = '22-Jan-2026'
date_format = pd.to_datetime(twenty_date) 
df = df[df['Created Date']>date_format]
df['Created Date'].count
df['Loan Number'] = df['Loan Number'].apply(lambda x:x.strip() if isinstance(x,str) else x)
save_path = r"C:\Users\ET0001301\Desktop\Python\FCL WHOLE DATA" #next tine choose this to append the data
yesterday_date = pd.Timestamp.now().date()
file_name = f'fcl_data-{yesterday_date}.xlsx'
import os
final_path = os.path.join(save_path,file_name)
 

df.columns.get_loc('Loan Number')
df.insert(29,'Len of PR',pd.NA)
df.columns
ecl_path = r"C:\Users\ET0001301\Desktop\Python\FCL WHOLE DATA\1.LAP_PORTFOLIO_JAN26-Final.xlsx"
ecl = pd.read_excel(ecl_path,sheet_name='Sheet1')
fcl_path = r"C:\Users\ET0001301\Desktop\Python\FCL WHOLE DATA\FCL Request - 12th Feb'26.xlsx"
fcl = pd.read_excel(fcl_path,sheet_name='Fresh Cases')
df = df[~df['Loan Number'].isin(fcl['Lan no.'])]
df['DPD BUCKET'] = df['Loan Number'].map(ecl.set_index('PROSPECTCODE')['DPD_BKT'])
df = df[df['DPD BUCKET'].isin(['A.Current','B.1-30'])]
df = df.drop_duplicates(subset=['Loan Number'])
# new file to append the data
fresh = df[['Loan Number']]
fresh = fresh.merge(ecl[['PROSPECTCODE','LEVIOSA ID','CUSTOMER_NAME','BRANCH','ME region ( Revised )','ZONE','POS in cr']],left_on='Loan Number',
                    right_on='PROSPECTCODE',how = 'left')
fresh.columns
df.columns
fresh = fresh.merge(df[['Loan Number','Created Date']],on = 'Loan Number',how = 'left')
fresh['Month'] = fresh['Created Date'].dt.strftime('%b')
fresh = fresh.merge(ecl[['PROSPECTCODE','DPD_BKT']],left_on='Loan Number',right_on='PROSPECTCODE',how = 'left')
fresh.columns
fresh = fresh.drop(columns = ['PROSPECTCODE_x','PROSPECTCODE_y'])
fresh.columns
fresh = fresh.rename(columns = {'Loan Number':'Lan no.','LEVIOSA ID':'Ref number',
                                'CUSTOMER_NAME':'Customer Name','BRANCH':'Branch',
                                'ME region ( Revised )':'Region 2','ZONE':'Zone',
                                'POS in cr':'New POS in cr','Created Date':'In Date',
                                'Month':'Month','DPD_BKT':'DPD'})
fresh
appended_data = pd.concat([fcl,fresh],ignore_index=True)
appended_data
appended_data = appended_data.drop(columns=['Retained Remark','Retained Stage','Branch remark ','Branch stage','Unnamed: 14','Unnamed: 15'])
appended_data['New POS in cr']=appended_data['New POS in cr'].apply(lambda x:round(x,2) )
appended_data['Day'] = appended_data['In Date'].dt.strftime('%d-%b')
appended_data
appended_data['In Date'] = appended_data['In Date'].dt.date
appended_data
append_pivot = pd.pivot_table(appended_data,
                              index = 'Region 2',
                              columns = 'In Date',
                              values='New POS in cr',
                              aggfunc='sum',
                              margins = True,
                              margins_name='Grand Total',
                              fill_value=0
                              )
append_pivot = append_pivot.astype(float).round(2)
append_pivot
with pd.ExcelWriter(final_path,engine='openpyxl',mode = 'w') as writer:
    appended_data.to_excel(writer,sheet_name='Data',index=False)
    append_pivot.to_excel(writer,sheet_name='pivot')
from openpyxl import load_workbook
wb = load_workbook(final_path)
ws = wb['Data']
for cols in ws.columns:
    max_length = 0
    col_letter = cols[0].column_letter
    for cell in cols:
        if cell.value is not None:
            max_length = max(max_length,len(str(cell.value)))
    ws.column_dimensions[col_letter].width = max_length+3
wb.save(final_path)
wb.close()
