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
