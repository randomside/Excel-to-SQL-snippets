import pandas as pd

item_master = pd.read_excel(
    r"data/item_master.xlsx", index_col=False, header=1)
material_history = pd.read_excel(
    'data/Material Movement History.xlsx', index_col=False, header=1)

temp_item_master = item_master[['Material', 'Material Type', 'Procurement']]

step_one = pd.merge(material_history, temp_item_master,
                    on="Material", how="left")

step_one.rename(columns={
    'Reference': 'Order Category',
    'Unnamed: 13': 'Order Data',
    'Unnamed: 14': 'Master Category',
    'Unnamed: 15': 'Master Data',
    'Unnamed: 16': 'Remark'
}, inplace=True)

step_one = step_one.iloc[1:, :]

# pd.set_option('mode.chained_assignment', None)

step_one.loc[step_one['Quantity'] > 0, 'Input'] = step_one['Quantity']
step_one.loc[step_one['Quantity'] > 0, 'Output'] = 0
step_one.loc[step_one['Quantity'] <= 0, 'Input'] = 0
step_one.loc[step_one['Quantity'] <= 0, 'Output'] = -step_one['Quantity']

cols = step_one.columns.tolist()

cols = cols[:10] + [cols[-2], cols[-1]] + cols[10:-2]

step_one = step_one[cols]

step_one['Account code'] = [int(x)
                            for x in (step_one['Movement Type'].str[:3])]

cols = step_one.columns.tolist()
cols = cols[:7] + [cols[-1]] + cols[7: -1]
step_one = step_one[cols]

step_one.to_excel("output/SCM_RAW.xlsx")

columns_name = {
    'H': 'Account code',
    'p': 'Order Category',
    'W': 'Material Type',
    'X': 'Procurement'
}

nhap_vao = [None] * 7
nhap_vao[0] = ((step_one['Account code'] == 101) & (
    step_one['Order Category'] == 'Purchase Order'))
nhap_vao[1] = (step_one['Account code'] == 602)
nhap_vao[2] = (step_one['Account code'] == 623)
nhap_vao[3] = (step_one['Account code'] == 701)
nhap_vao[4] = (step_one['Account code'] == 720)
nhap_vao[5] = (step_one['Account code'] == 801)
nhap_vao[6] = (step_one['Account code'] == 809)

xuat_ra = [None] * 9
xuat_ra[0] = ((step_one['Account code'] == 102) & (
    step_one['Order Category'] == 'Purchase Order'))
xuat_ra[1] = (step_one['Account code'] == 201)
xuat_ra[2] = (step_one['Account code'] == 261)
xuat_ra[3] = (step_one['Account code'] == 601)
xuat_ra[4] = (step_one['Account code'] == 609)
xuat_ra[5] = (step_one['Account code'] == 702)
xuat_ra[6] = (step_one['Account code'] == 712)
xuat_ra[7] = (step_one['Account code'] == 721)
xuat_ra[8] = (step_one['Account code'] == 803)

thuyen_chuyen = [None] * 2
thuyen_chuyen[0] = (step_one['Account code'].between(300, 400))
thuyen_chuyen[1] = (step_one['Account code'] == 401)


def condition_to_data(arr):
    material_type_order = ['ROH', 'HAWA', 'HALB', 'FERT']
    material_order_index = dict(
        zip(material_type_order, range(len(material_type_order))))
    for i in range(len(arr)):
        arr[i] = step_one.loc[arr[i]]
        arr[i]['Tm_Rank'] = arr[i]['Material Type'].map(material_order_index)
        arr[i].sort_values(['Tm_Rank', 'Procurement'], inplace=True)
        arr[i].drop('Tm_Rank', 1, inplace=True)


condition_to_data(nhap_vao)
condition_to_data(xuat_ra)
condition_to_data(thuyen_chuyen)

step_two = pd.concat(nhap_vao + xuat_ra + thuyen_chuyen)

condition_to_delete = ~(step_two['Account code'].isin([101, 102])) & (
    ((step_two['Material Type'] == 'ROH') & (step_two['Procurement'] == 'E')) |
    ((step_two['Material Type'] == 'HAWA') & (step_two['Procurement'] == 'E')) |
    ((step_two['Material Type'] == 'HALB') & (step_two['Procurement'].isin(['E', 'X']))) |
    ((step_two['Material Type'] == 'FERT') &
     (step_two['Procurement'].isin(['E', 'X'])))
)
step_two.drop(step_two[condition_to_delete].index, inplace=True)

SCM_RAW = step_two

nhap_vao = [None] * 2
nhap_vao[0] = (
    (step_one['Account code'] == 101) &
    (step_one['Order Category'] == 'Purchase Order') &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

nhap_vao[1] = (
    (step_one['Account code'].isin([602, 623, 701, 720, 801, 809])) &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra = [None] * 3
xuat_ra[0] = (
    (step_one['Account code'] == 102) &
    (step_one['Order Category'] == 'Purchase Order') &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra[1] = (
    (step_one['Account code'].isin([201, 601, 609, 702, 712, 721, 803])) &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra[2] = (
    (step_one['Account code'] == 261) &
    (
        ((step_one['Material Type'] == 'HALB') &
         (step_one['Procurement'].isin(['E', 'X']))) |
        ((step_one['Material Type'] == 'FERT') &
         (step_one['Procurement'].isin(['E', 'X'])))
    )
)

thuyen_chuyen = [None]
thuyen_chuyen[0] = (
    (step_one['Account code'].between(300, 402)) &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

condition_to_data(nhap_vao)
condition_to_data(xuat_ra)
condition_to_data(thuyen_chuyen)

SCM_WIP = pd.concat(nhap_vao + xuat_ra + thuyen_chuyen)

SCM_WIP.loc[(
    (SCM_WIP['Account code'] == 261) &
    (SCM_WIP['Material Type'] == 'FERT') &
    (SCM_WIP['Location'].str.startswith('RW'))
), 'Location'] = None

nhap_vao = [None] * 2
nhap_vao[0] = (
    (step_one['Account code'] == 101) &
    (step_one['Order Category'] == 'Production Order') &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

nhap_vao[1] = (
    (step_one['Account code'].isin([602, 623, 701, 720, 801, 809])) &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra = [None] * 2
xuat_ra[0] = (
    (step_one['Account code'] == 102) &
    (step_one['Order Category'] == 'Production Order') &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra[1] = (
    (step_one['Account code'].isin([201, 601, 609, 702, 712, 721, 803])) &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

thuyen_chuyen = [None]
thuyen_chuyen[0] = (
    (step_one['Account code'].between(300, 402)) &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

condition_to_data(nhap_vao)
condition_to_data(xuat_ra)
condition_to_data(thuyen_chuyen)

SCM_FG = pd.concat(nhap_vao + xuat_ra + thuyen_chuyen)

print('RAW', SCM_RAW.shape)
print('FG', SCM_FG.shape)
print('WIP', SCM_WIP.shape)

with pd.ExcelWriter('output/before_report.xlsx') as writer:
    SCM_RAW.to_excel(writer, sheet_name='RAW')
    SCM_FG.to_excel(writer, sheet_name='FG')
    SCM_WIP.to_excel(writer, sheet_name='WIP')

print('DONE IT')
