import pandas as pd
import numpy as np
import calendar
import datetime
import openpyxl

Bill = pd.read_excel("combined.xlsx",sheet_name="Bill")
Transaction = pd.read_excel("combined.xlsx",sheet_name="Transaction")

############### BILL DATA ##############
# Formatting 
Bill["Due Date"] = pd.to_datetime(Bill["Due Date"])
Bill["Contract Date"] = pd.to_datetime(Bill["Contract Date"])
Bill.sort_values(["CL Contract: CL Contract ID","Contract Date","Due Date"],inplace=True)
Bill["Year"] = Bill["Due Date"].apply(lambda x: x.year)
Bill["Month"] = Bill["Due Date"].apply(lambda x: x.month)
Bill["Month"] = Bill["Month"].apply(lambda x: calendar.month_name[x])
Bill["Time"] = Bill["Month"].map(str) + "_" + Bill["Year"].map(str)
Bill["UID"] = Bill["CL Contract: CL Contract ID"].map(str) +"-"+ Bill["Time"].map(str)  

############### Transaction  DATA ##############
Transaction["Transaction Date"] = pd.to_datetime(Transaction["Transaction Date"])
Transaction["Year"] = Transaction["Transaction Date"].apply(lambda x: x.year)
Transaction["Month"] = Transaction["Transaction Date"].apply(lambda x: x.month)
Transaction["Month"] = Transaction["Month"].apply(lambda x: calendar.month_name[x])
Transaction["Time"] = Transaction["Month"].map(str) + "_" + Transaction["Year"].map(str)
Transaction["UID"] = Transaction["Loan Account: CL Contract ID"].map(str) +"-"+ Transaction["Time"].map(str)
Transaction = Transaction.groupby(["UID"],as_index=False).sum()

##############################################
##############################################
# Merging Transaction and Bill data.
master = pd.merge(left=Bill,right=Transaction,how="left",left_on="UID",right_on="UID")
master["Current Payment Amount"] = np.ceil(master["Current Payment Amount"])
master["Transaction Amount"] = np.ceil(master["Transaction Amount"])

##############################################
##############################################

# Split it into Buckets.

year_list = list(range(2016,2019))

# print(year_list)
# September Data Filtering Manually for now.
# master.to_csv("master.csv",index=False)

def monthly_vintage(month_data,contract_list):
	col_list = ["CL Contract: CL Contract ID","Contract Date","Due Date","Current Payment Amount","Year_x","Month","Time","Transaction Amount"]
	dummy_dict=dict()

	iter=0

	for cid in contract_list:
		dummy_dict["cust_"+str(iter)] = month_data.loc[month_data["CL Contract: CL Contract ID"]==cid,col_list]
		dummy_dict["cust_"+str(iter)]["Cumulative_payment_amount"] = np.cumsum(dummy_dict["cust_"+str(iter)]["Current Payment Amount"])
		dummy_dict["cust_"+str(iter)]["Cumulative_Transaction_amount"] = np.cumsum(dummy_dict["cust_"+str(iter)]["Transaction Amount"])
		dummy_dict["cust_"+str(iter)]["Cumulative_Transaction_amount"].fillna(method='ffill',inplace=True)
		dummy_dict["cust_"+str(iter)]["Cumulative_Transaction_amount"].fillna(0,inplace=True)
		dummy_dict["cust_"+str(iter)]["Monthly_delinquent"] = dummy_dict["cust_"+str(iter)]["Cumulative_payment_amount"] - dummy_dict["cust_"+str(iter)]["Cumulative_Transaction_amount"]
		dummy_dict["cust_"+str(iter)]["DPD"] = np.ceil(dummy_dict["cust_"+str(iter)]["Monthly_delinquent"]/dummy_dict["cust_"+str(iter)]["Current Payment Amount"])
		dummy_dict["cust_"+str(iter)]["MOB"] = list(range(1,len(dummy_dict["cust_"+str(iter)]["Due Date"])+1))
		dummy_dict["cust_"+str(iter)].loc[dummy_dict["cust_"+str(iter)]["DPD"]>0,"30+DPD"]=1
		dummy_dict["cust_"+str(iter)].loc[dummy_dict["cust_"+str(iter)]["DPD"]>1,"60+DPD"]=1
		dummy_dict["cust_"+str(iter)].loc[dummy_dict["cust_"+str(iter)]["DPD"]>2,"90+DPD"]=1
		dummy_dict["cust_"+str(iter)].loc[dummy_dict["cust_"+str(iter)]["DPD"]>3,"90+DPD"]=1
		dummy_dict["cust_"+str(iter)].loc[dummy_dict["cust_"+str(iter)]["DPD"]>4,"120+DPD"]=1
		dummy_dict["cust_"+str(iter)].loc[dummy_dict["cust_"+str(iter)]["DPD"]>5,"150+DPD"]=1
		iter=iter+1

	cat = pd.concat([dummy_dict["cust_"+str(i)] for i in range(0,iter)])

	cat_dpd = cat[["MOB","30+DPD","60+DPD","90+DPD","120+DPD","150+DPD"]]
	cat_dpd = cat_dpd.groupby(["MOB"],as_index=False).sum()
	cat_dpd["Total Number of Loans"]=len(contract_list)
	return cat_dpd


vintage_dict=dict()
vintage_contract_list = dict()
vintage_final = dict()

time_iteration = pd.read_excel("Date_iteration.xlsx",sheet_name="Sheet1")
time_iteration["Start_date"] = pd.to_datetime(time_iteration["Start_date"])
time_iteration["End_date"] = pd.to_datetime(time_iteration["End_date"])

writer = pd.ExcelWriter('output.xlsx')

for i in range(0,len(time_iteration["Iteration"])):
	vintage_dict["vintage_"+str(i)] = master.loc[(master["Contract Date"]<time_iteration.iloc[i,2]),]
	vintage_dict["vintage_"+str(i)] = vintage_dict["vintage_"+str(i)].loc[(vintage_dict["vintage_"+str(i)]["Contract Date"]>=time_iteration.iloc[i,1]),]
	vintage_contract_list["vintage_"+str(i)] = sorted(list(set(vintage_dict["vintage_"+str(i)]["CL Contract: CL Contract ID"])))
	vintage_final["vintage_"+str(i)] = monthly_vintage(vintage_dict["vintage_"+str(i)],vintage_contract_list["vintage_"+str(i)])
	vintage_final["vintage_"+str(i)].to_excel(writer,sheet_name=time_iteration.iloc[i,0])
	writer.save()

# vintage_final["vintage_"+str(i)].to_excel(writer,1,startrow=1,startcol=1)
# writer.save()
# vintage_dict["vintage_"+str(i)] = master.loc[(master["Contract Date"]>time_iteration.iloc[0,1] & master["Contract Date"]<time_iteration.iloc[0,2]) == True,]