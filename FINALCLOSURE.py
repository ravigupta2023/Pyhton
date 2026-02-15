

# -------------------- READ DATA --------------------
# path = r"C:\Users\ET0001301\Pictures\DailyClosurePivot.xlsm"
df = pd.read_excel(shutil_final, sheet_name='Sheet1')

data = df[
    (df['ProductCheck'] != 'Rest') &
    (df['REASON CODE'].isin(['CANCELLATION', 'NORMAL FORECLOSURE']))
]

data['IN_CR'] = pd.to_numeric(data['IN_CR'], errors='coerce')

# -------------------- BASE PIVOT (REGION LEVEL) --------------------
base = (
    data
    .groupby(['Zone', 'Region2', 'REASON CODE'])['IN_CR']
    .sum()
    .unstack(fill_value=0)
    .reset_index()
)

base = base.rename(columns={'NORMAL FORECLOSURE': 'CLOSURE'})
base['Grand Total'] = base['CANCELLATION'] + base['CLOSURE']

# -------------------- TARGET (MAX CLOSURE) --------------------
target = {
    'Assam':0.8,'WB':0.0,'Bihar':2.2,
    'Delhi':2.9,'Haryana':0.7,'PCH':0.2,'Rajasthan':0.6,'Uttarakhand':1.1,
    'UP East':1.2,'UP West':1.6,'Varanasi':0.7,
    'Bangalore':3.2,'ROK':0.6,
    'Telangana 1':1.1,'Telangana 2':2.2,
    'AP':1.5,'TN':0.9,
    'Mumbai':0.6,'ROM':0.3,'ROM 2':0.5,
    'Gujarat':0.4,'MP':1.7
}

base['Max Closure'] = base['Region2'].map(target).fillna(0)

# -------------------- % CALC --------------------
base['Max Closure %'] = (
    base['Grand Total']
    .div(base['Max Closure']).mul(100)
    .replace([float('inf'), -float('inf')], 0)
    .round(0).astype(str) + '%'
)

# -------------------- SUBTOTAL FUNCTION --------------------
def add_total(df, name):
    row = df[['CANCELLATION','CLOSURE','Grand Total','Max Closure']].sum().to_frame().T
    row['Max Closure %'] = (
        row['Grand Total'].div(row['Max Closure']).mul(100)
    ).fillna(0).round(0).astype(int).astype(str) + '%'
    row['Zone'] = ''
    row['Region2'] = name
    return pd.concat([df, row], ignore_index=True)

final = []

# -------------------- EAST --------------------
east = base[base['Zone'] == 'East']
east = add_total(east, 'East Total')
final.append(east)

# -------------------- NORTH (WITH DELHI & UP TOTALS) --------------------
north = base[base['Zone'] == 'North']

delhi_states = ['Delhi','Haryana','PCH','Rajasthan','Uttarakhand']
delhi = north[north['Region2'].isin(delhi_states)]
delhi_total = add_total(delhi, 'Delhi Total')

up = north[north['Region2'].isin(['UP East','UP West','Varanasi'])]
up_tatal = add_total(up, 'UP Total')

north_concat = pd.concat([delhi_total,up_tatal],ignore_index=True)

north_final = north[north['Region2'].isin(['UP East','UP West','Varanasi',
                                           'Delhi','Haryana','PCH','Rajasthan','Uttarakhand'])]
north_total = add_total(north_final,'North Total')

north_total1 = pd.concat([north_concat,north_total.tail(1)],ignore_index=True)

final.append(north_total1)

# -------------------- SOUTH --------------------

south = base[base['Zone']=='South']
karnataka_states = ['Bangalore','ROK','Mysore']
telangana_states = ['Telangana 1','Telangana 2']

karnatala_total = south[south['Region2'].isin(karnataka_states)]
telangana_total = south[south['Region2'].isin(telangana_states)]

karnatala_total1 = add_total(karnatala_total,'karnataka Total')
telangana_total1 = add_total(telangana_total,'Telangana Total')

remain_south = ['TN',"AP"]
remain_south1 = south[south['Region2'].isin(remain_south)]

karnataka_add = pd.concat([karnatala_total,karnatala_total1.tail(1),telangana_total,telangana_total1.tail(1),remain_south1],
                          ignore_index=True)
final_south = south[south['Region2'].isin(['Telangana 1','Telangana 2','TN','AP','Bangalore','ROK','Mysore'])]
final_south_total = add_total(final_south,'South Total')

south_total = pd.concat([karnataka_add,final_south_total.tail(1)],ignore_index=True)
final.append(south_total) 

# -------------------- WEST --------------------


# -------------------- WEST (FINAL FIX) --------------------
west = base[base['Zone'] == 'West']

# Maharashtra details
maha_detail = west[west['Region2'].isin(['Mumbai','ROM','ROM 2'])]
maha_total = add_total(maha_detail, 'Maharashtra Total')

# MP & Gujarat
west_rest = west[west['Region2'].isin(['MP','Gujarat'])]

# Display block (details + Maharashtra Total)
west_block = pd.concat(
    [maha_detail, maha_total.tail(1), west_rest],
    ignore_index=True
)

# West Total → calculated ONLY from base rows
west_total_base = west[
    west['Region2'].isin(['Mumbai','ROM','ROM 2','MP','Gujarat'])
]
west_total = add_total(west_total_base, 'West Total')

# Append ONLY the total row
west_block = pd.concat(
    [west_block, west_total.tail(1)],
    ignore_index=True
)

final.append(west_block)

# -------------------- FINAL MERGE --------------------
final_df = pd.concat(final, ignore_index=True)

# -------------------- GRAND TOTAL --------------------
grand = base[['CANCELLATION','CLOSURE','Grand Total','Max Closure']].sum().to_frame().T
grand['Max Closure %'] = (
    grand['Grand Total'].div(grand['Max Closure']).mul(100)
).round(0).astype(int).astype(str) + '%'
grand['Zone'] = ''
grand['Region2'] = 'Grand Total'

final_df = pd.concat([final_df, grand], ignore_index=True)


