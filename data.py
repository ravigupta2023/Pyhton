import pandas as pd

path  = r"C:\Users\Ravi Gupta\OneDrive\Desktop\ManpowerME & LAG.xlsm"

df = pd.read_excel(path)

# print(df.head())
new_data  = df.drop(df.columns[[0,7,12,14,17,18,20,21,22,23]],axis = 1)

new_data.rename(columns = {'ReportingManager_Code':'RM_CODE_1'},inplace = True)

# print(new_data.columns)

rm_code_1 = ['Employee_Code','Employee_Name','Role']
base_info = df[rm_code_1]
merge_data = pd.merge(new_data,base_info,left_on = 'RM_CODE_1',right_on = 'Employee_Code',how = 'left')
merge_data.rename(columns={'Employee_Name_y':'RM_NAME_1','Role_y':'RM_DESIG_1'},inplace = True)
merge_data = merge_data.drop(columns=['Employee_Code_y'])

rm_code_2 = ['Employee_Code','ReportingManager_Code']
base_info2 = df[rm_code_2]
merge_data = pd.merge(merge_data,base_info2,left_on = 'RM_CODE_1',right_on = 'Employee_Code',how = 'left')
merge_data.drop(columns = ['Employee_Code'],inplace = True)
merge_data.rename(columns = {'Employee_Code_x':'Employee_Code','ReportingManager_Code':'RM_CODE_2'},inplace = True)

rm_code_3 = ['Employee_Code','Employee_Name','Role']
base_info3 = df[rm_code_3]
merge_data = pd.merge(merge_data,base_info3,left_on = 'RM_CODE_2',right_on = 'Employee_Code',how = 'left')
merge_data.rename(columns ={'Employee_Name':'RM_NAME_2','Role':'RM_DESIG_2'},inplace = True)
merge_data.drop(columns = ['Employee_Code_y'],inplace = True)
merge_data.rename(columns={'Employee_Code_x':'Employee_Code','Employee_Name_x':'Employee_Name'},inplace = True)  

merge_data.fillna('#N/A',inplace = True)


import os

new_file = r"C:\Users\Ravi Gupta\OneDrive\Desktop"
name = 'mapped_file.xlsx'

full_path = os.path.join(new_file,name)

merge_data.to_excel(full_path,index = False)     