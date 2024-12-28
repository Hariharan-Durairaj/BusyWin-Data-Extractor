from busywin_tools import busywin_transactions
data=busywin_transactions()
df=data.get_data()
print(df)
# to export the transaction data
#data.export_data()