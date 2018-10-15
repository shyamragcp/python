import pandas as pd
import numpy as np
import calendar
import datetime
import openpyxl
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

pd.set_option('chained',None)

Bill = pd.read_excel("combined.xlsx",sheet_name="Bill")
Transaction = pd.read_excel("combined.xlsx",sheet_name="Transaction")
Disbursal = pd.read_excel("combined.xlsx",sheet_name="Disbursal")
date_iteration = pd.read_excel("Date_iteration.xlsx",sheet_name="Sheet1")

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

unique_due_list = sorted(Bill["Due Date"].unique())
due_list_df = pd.DataFrame()
due_list_df["Due Date"] = unique_due_list
due_list_df["Year"] = due_list_df["Due Date"].apply(lambda x: x.year)
due_list_df["Month"] = due_list_df["Due Date"].apply(lambda x: x.month)
due_list_df["Month"] = due_list_df["Month"].apply(lambda x: calendar.month_name[x])
due_list_df["Time"] = due_list_df["Month"].map(str)+"_"+due_list_df["Year"].map(str)

Bill_unique = Bill.drop_duplicates(subset=["CL Contract: CL Contract ID"],keep="last")
Bill_unique = Bill_unique[["CL Contract: CL Contract ID","Current Payment Amount","Contract Date"]]

############### Transaction  DATA ##############
Transaction["Transaction Date"] = pd.to_datetime(Transaction["Transaction Date"])
Transaction["Year"] = Transaction["Transaction Date"].apply(lambda x: x.year)
Transaction["Month"] = Transaction["Transaction Date"].apply(lambda x: x.month)
Transaction["Month"] = Transaction["Month"].apply(lambda x: calendar.month_name[x])
Transaction["Time"] = Transaction["Month"].map(str) + "_" + Transaction["Year"].map(str)
Transaction["UID"] = Transaction["Loan Account: CL Contract ID"].map(str) +"-"+ Transaction["Time"].map(str)
Transaction["Transaction Amount"] = Transaction["Transaction Amount"] - Transaction["Fees"]
Transaction.sort_values(["Loan Account: CL Contract ID","Contract Date","Transaction Date"],inplace=True)

Bill_sample = Bill[["CL Contract: CL Contract ID","Contract Date"]]
Bill_sample.drop_duplicates(subset="CL Contract: CL Contract ID",keep="last",inplace=True) 
Bill_cumsum = Bill[["CL Contract: CL Contract ID","Contract Date","Due Date","Current Payment Amount"]]
Bill_cumsum["Cumilative Payment Amount"] = Bill_cumsum.groupby(["CL Contract: CL Contract ID"]).cumsum()

Transaction_cum = Transaction[["Loan Account: CL Contract ID","Contract Date","Transaction Date","Transaction Amount"]]
Transaction_cum["Cumilative Transaction Amount"] = Transaction_cum.groupby(["Loan Account: CL Contract ID"]).cumsum()
principal_cum = Transaction[["Loan Account: CL Contract ID","Contract Date","Transaction Date","Principal"]]
principal_cum["Principal Cum"] = principal_cum.groupby(["Loan Account: CL Contract ID"]).cumsum()

#########################
#### Main Body ## 
#########################

test_writer = pd.ExcelWriter('result.xlsx')
portfolio_dict = dict()
final_df = pd.DataFrame(index=["Total","Current","<30 DPD","31-60 DPD","61-90 DPD","91-120 DPD","120+ DPD"],columns  = date_iteration["Iteration"])
final_Roll = dict()

f=open("out.txt","w")

def segmentation(loop,sheet_name,i):

	sample_d = Bill_sample.loc[Bill_sample["Contract Date"]<=pd.to_datetime(loop),]
	sample_d = sample_d.drop_duplicates(subset=["CL Contract: CL Contract ID"],keep="last")
	Bill_d = Bill_cumsum.loc[Bill_cumsum["Due Date"]<pd.to_datetime(loop),]
	Bill_d = Bill_d.drop_duplicates(subset=["CL Contract: CL Contract ID"],keep="last")
	Transaction_d = Transaction_cum.loc[Transaction_cum["Transaction Date"]<=pd.to_datetime(loop),]
	Transaction_d.sort_values(["Loan Account: CL Contract ID","Transaction Date"],inplace=True)
	Transaction_d = Transaction_d.drop_duplicates(subset=["Loan Account: CL Contract ID"],keep="last")

	principal_d = principal_cum.loc[principal_cum["Transaction Date"]<=pd.to_datetime(loop),]
	principal_d.sort_values(["Loan Account: CL Contract ID","Transaction Date"],inplace=True)
	principal_d = principal_d.drop_duplicates(subset=["Loan Account: CL Contract ID"],keep="last")

	Transaction_d = pd.merge(left=Transaction_d,right=principal_d[["Loan Account: CL Contract ID","Principal Cum"]],how="left",left_on="Loan Account: CL Contract ID",right_on="Loan Account: CL Contract ID")

	merged = pd.merge(left=sample_d,right=Bill_d,how="left",left_on="CL Contract: CL Contract ID",right_on="CL Contract: CL Contract ID")
	merged = pd.merge(left=merged,right=Transaction_d,how="left",left_on="CL Contract: CL Contract ID",right_on="Loan Account: CL Contract ID")
	Transaction_E = Transaction_cum.loc[Transaction_cum["Contract Date"]<=pd.to_datetime(loop),]
	Transaction_E = Transaction_E.drop_duplicates(subset=["Loan Account: CL Contract ID"],keep="last")

	Transaction_E = pd.merge(left=Transaction_E,right=principal_d[["Loan Account: CL Contract ID","Principal Cum"]],how="left",left_on="Loan Account: CL Contract ID",right_on="Loan Account: CL Contract ID")

	early_settlement = list(set(Transaction_E["Loan Account: CL Contract ID"]) - set(merged["CL Contract: CL Contract ID"]))
	if len(early_settlement) >0:
		early_df = Transaction_E.loc[Transaction_E["Loan Account: CL Contract ID"].isin(early_settlement) == True,]
		early_df = early_df.loc[early_df["Contract Date"] != early_df["Transaction Date"],]
		if len(early_df)>0:
			early_df.rename(columns={"Cumilative Transaction Amount":"Full_Amount"},inplace=True)
			merged = pd.concat([merged,early_df],ignore_index=True,axis=0)
	merged.loc[merged["CL Contract: CL Contract ID"].isna() == True,"CL Contract: CL Contract ID"] = merged.loc[merged["CL Contract: CL Contract ID"].isna() == True,"Loan Account: CL Contract ID"]
	merged.loc[merged["Contract Date"].isna() == True,"Contract Date"] = merged.loc[merged["Contract Date"].isna() == True,"Contract Date_x"]
	merged["Cumilative Transaction Amount"].fillna(0,inplace=True)
	merged["DPD"] = np.ceil(round((merged["Cumilative Payment Amount"] - merged["Cumilative Transaction Amount"]),2)/merged["Current Payment Amount"])
	merged.loc[merged["DPD"].isna() == True,"DPD Status"]="Current"
	merged.loc[merged["DPD"] <= 0,"DPD Status"]="Current"
	merged.loc[merged["DPD"] == 1,"DPD Status"]="< 30 DPD"
	merged.loc[merged["DPD"] == 2,"DPD Status"]="31-60 DPD"
	merged.loc[merged["DPD"] == 3,"DPD Status"]="61-90 DPD"
	merged.loc[merged["DPD"] == 4,"DPD Status"]="91-120 DPD"
	merged.loc[merged["DPD"] >= 5,"DPD Status"]="120+ DPD"
	portfolio_dict[sheet_name] = merged
	final_df.iloc[0,i] = len(merged)
	final_df.iloc[1,i] = len(merged.loc[merged["DPD Status"] == "Current","DPD"])
	final_df.iloc[2,i] = len(merged.loc[merged["DPD Status"] == "< 30 DPD","DPD"])
	final_df.iloc[3,i] = len(merged.loc[merged["DPD Status"] == "31-60 DPD","DPD"])
	final_df.iloc[4,i] = len(merged.loc[merged["DPD Status"] == "61-90 DPD","DPD"])
	final_df.iloc[5,i] = len(merged.loc[merged["DPD Status"] == "91-120 DPD","DPD"])
	final_df.iloc[6,i] = len(merged.loc[merged["DPD Status"] == "120+ DPD","DPD"])
	if "Full_Amount" in merged.columns:
		merged.loc[merged["CL Contract: CL Contract ID"].isin(early_settlement),"Transaction Amount"] = merged.loc[merged["CL Contract: CL Contract ID"].isin(early_settlement),"Full_Amount"]
	merged.rename(columns={"DPD Status":"DPD Status"+sheet_name},inplace=True)
	merged = pd.merge(left=merged,right=Disbursal,how="left",left_on="CL Contract: CL Contract ID",right_on="CL Contract ID")
	merged.rename(columns={"Disbursal Transaction Amount":"Disbursal"+sheet_name},inplace=True)
	merged.rename(columns={"Principal Cum":"Principal_paid"+sheet_name},inplace=True)
	merged.to_excel(test_writer,sheet_name=sheet_name)
	test_writer.save()
	if i>0:
		merged.loc[merged["Transaction Date"]<=date_iteration["Report_Date"][i-1],"Transaction Amount"] = 0
	merged.rename(columns={"Transaction Amount":"Transaction Amount"+sheet_name},inplace=True)

	### Different DPDS
	DPD_30 = merged.copy()
	DPD_30.loc[DPD_30["DPD Status"+sheet_name].isin(["< 30 DPD","Current"])==True,"Principal_paid"+sheet_name] = 0
	DPD_30.loc[DPD_30["DPD Status"+sheet_name].isin(["< 30 DPD","Current"])==True,"Disbursal"+sheet_name] = 0
	DPD_30["Principal_outstanding"+sheet_name] = round(DPD_30["Disbursal"+sheet_name] - DPD_30["Principal_paid"+sheet_name],1)
	DPD_60 = merged.copy()
	DPD_60.loc[DPD_60["DPD Status"+sheet_name].isin(["31-60 DPD","< 30 DPD","Current"])==True,"Principal_paid"+sheet_name] = 0
	DPD_60.loc[DPD_60["DPD Status"+sheet_name].isin(["31-60 DPD","< 30 DPD","Current"])==True,"Disbursal"+sheet_name] = 0
	DPD_60["Principal_outstanding"+sheet_name] = round(DPD_60["Disbursal"+sheet_name] - DPD_60["Principal_paid"+sheet_name],1)
	DPD_90 = merged.copy()
	DPD_90.loc[DPD_90["DPD Status"+sheet_name].isin(["31-60 DPD","61-90 DPD","< 30 DPD","Current"])==True,"Principal_paid"+sheet_name] = 0
	DPD_90.loc[DPD_90["DPD Status"+sheet_name].isin(["31-60 DPD","61-90 DPD","< 30 DPD","Current"])==True,"Disbursal"+sheet_name] = 0
	DPD_90["Principal_outstanding"+sheet_name] = round(DPD_90["Disbursal"+sheet_name] - DPD_90["Principal_paid"+sheet_name],1)
	DPD_120 = merged.copy()
	DPD_120.loc[DPD_120["DPD Status"+sheet_name].isin(["31-60 DPD","61-90 DPD","91-120 DPD","< 30 DPD","Current"])==True,"Principal_paid"+sheet_name] = 0
	DPD_120.loc[DPD_120["DPD Status"+sheet_name].isin(["31-60 DPD","61-90 DPD","91-120 DPD","< 30 DPD","Current"])==True,"Disbursal"+sheet_name] = 0
	DPD_120["Principal_outstanding"+sheet_name] = round(DPD_120["Disbursal"+sheet_name] - DPD_120["Principal_paid"+sheet_name],1)

	#### return Items
	R_1 = merged[["CL Contract: CL Contract ID","DPD Status"+sheet_name]]
	R_2 = DPD_30[["CL Contract: CL Contract ID","Principal_outstanding"+sheet_name]]
	R_3 = DPD_60[["CL Contract: CL Contract ID","Principal_outstanding"+sheet_name]]
	R_4 = DPD_90[["CL Contract: CL Contract ID","Principal_outstanding"+sheet_name]]
	R_5 = DPD_120[["CL Contract: CL Contract ID","Principal_outstanding"+sheet_name]]
	return R_1,R_2,R_3,R_4,R_5

i=0
for loop in date_iteration["Report_Date"]:
	sheet_name = date_iteration["Iteration"][i]
	merged,vintage_30,vintage_60,vintage_90,vintage_120 = segmentation(loop,sheet_name,i)
	if i==0:
		final_Roll = merged
		vin30_width = vintage_30
		vin60_width = vintage_60
		vin90_width = vintage_90
		vin120_width = vintage_120
	else:
		final_Roll = pd.merge(left=merged,right=final_Roll,how="left",left_on="CL Contract: CL Contract ID",right_on="CL Contract: CL Contract ID")
		vin30_width = pd.merge(left=vintage_30,right=vin30_width,how="left",left_on="CL Contract: CL Contract ID",right_on="CL Contract: CL Contract ID")
		vin60_width = pd.merge(left=vintage_60,right=vin60_width,how="left",left_on="CL Contract: CL Contract ID",right_on="CL Contract: CL Contract ID")
		vin90_width = pd.merge(left=vintage_90,right=vin90_width,how="left",left_on="CL Contract: CL Contract ID",right_on="CL Contract: CL Contract ID")
		vin120_width = pd.merge(left=vintage_120,right=vin120_width,how="left",left_on="CL Contract: CL Contract ID",right_on="CL Contract: CL Contract ID")
	i=i+1

final_df.to_excel(test_writer,sheet_name="CIR")
test_writer.save()

#### Pivot Tables Swap in Swap Out.
swap_writer = pd.ExcelWriter('roll_matrix.xlsx')
final_Roll.to_excel(swap_writer,sheet_name="final_roll")
swap_writer.save()

#########################################################################################################################################################
#########################################################################################################################################################
#########################################################################################################################################################
#### Writting Vintage Input -- Should Use this Output as the input to vintage.py
vin_writer = pd.ExcelWriter("Vintage_input.xlsx")
def vin_save(df,name):
	df.to_excel(vin_writer,sheet_name=name)
	vin_writer.save()

vin_save(final_Roll,"Count_Vin")
vin_save(vin30_width,"vin30")
vin_save(vin60_width,"vin60")
vin_save(vin90_width,"vin90")
vin_save(vin120_width,"vin120")

# final_Roll.to_csv("Vintage_input.csv",index=False)
# vintage_width.to_csv("Vintage_Amount.csv",index=False)

roll_single = pd.ExcelWriter("Roll_B_F.xlsx")
row_jump = 0
roll_current = pd.DataFrame()
roll_1_30 = pd.DataFrame()
roll_30_60 = pd.DataFrame()
roll_60_90 = pd.DataFrame()
roll_90_120 = pd.DataFrame()
roll_120 = pd.DataFrame()

for colmn in range(1,len(final_Roll.columns)-1):
	sheet_name = final_Roll.columns[colmn]
	final_Roll_new = final_Roll.loc[final_Roll[final_Roll.columns[colmn]].isna()==False,]
	final_Roll_new.fillna("New",inplace=True)
	pivot_roll = pd.pivot_table(final_Roll_new,index=final_Roll_new.columns[colmn+1],columns=final_Roll_new.columns[colmn],values="CL Contract: CL Contract ID",aggfunc="count",fill_value =0)
	col_order = ["Current","< 30 DPD","31-60 DPD","61-90 DPD","91-120 DPD","120+ DPD","New"]
	pivot_roll = pd.DataFrame(pivot_roll)
	pivot_roll = pivot_roll.loc[col_order,col_order]
	pivot_roll.fillna(0,inplace=True)
	pivot_roll["Total"] = pivot_roll.sum(axis=1)
	pivot_roll.loc["Total"] = pivot_roll.sum()

	pivot_roll.loc["Current","Roll Back"] = 0
	for col_name in col_order[:-1]:
		pivot_roll.loc[col_name,"Stabilized"] = round(pivot_roll.loc[col_name,col_name]/pivot_roll.loc[col_name,"Total"]*100,1)

	lp = 1
	for col_name in col_order[:-1]:
		pivot_roll.loc[col_name,"Roll Forward"] = round(pivot_roll.loc[col_name,col_order[lp:]].sum()/pivot_roll.loc[col_name,"Total"]*100,1)
		pivot_roll.loc[col_name,"Roll Back"] = round(pivot_roll.loc[col_name,col_order[:lp-1]].sum()/pivot_roll.loc[col_name,"Total"]*100,1)
		lp = lp + 1
	
	pivot_roll.loc["120+ DPD","Roll Forward"] = 0

	pivot_roll.to_excel(swap_writer,sheet_name = final_Roll_new.columns[colmn])
	swap_writer.save()

	pivot_roll.to_excel(roll_single,sheet_name="Swap",startrow=10*row_jump+1,startcol=1)
	roll_single.save()

	pivot_roll.insert(0,"Col_index",final_Roll_new.columns[colmn])
	col_order2 = ["Col_index","Current","< 30 DPD","31-60 DPD","61-90 DPD","91-120 DPD","120+ DPD","New","Total","Roll Back","Stabilized","Roll Forward"]
	roll_current=roll_current.append(pivot_roll.loc["Current",col_order2])
	roll_1_30=roll_1_30.append(pivot_roll.loc["< 30 DPD",col_order2])
	roll_30_60=roll_30_60.append(pivot_roll.loc["31-60 DPD",col_order2])
	roll_60_90=roll_60_90.append(pivot_roll.loc["61-90 DPD",col_order2])
	roll_90_120=roll_90_120.append(pivot_roll.loc["91-120 DPD",col_order2])
	roll_120=roll_120.append(pivot_roll.loc["120+ DPD",col_order2])

	# roll_current.reset_index()
	roll_current = roll_current[col_order2]
	# # roll_1_30.reset_index()
	roll_1_30 = roll_1_30[col_order2]
	# # roll_30_60.reset_index()
	roll_30_60 = roll_30_60[col_order2]
	# # roll_60_90.reset_index()
	roll_60_90 = roll_60_90[col_order2]
	# # roll_90_120.reset_index()
	roll_90_120 = roll_90_120[col_order2]
	# # roll_120.reset_index()
	roll_120 = roll_120[col_order2]


	roll_current.to_excel(roll_single,sheet_name="cluster",startrow=1,startcol=1)
	roll_1_30.to_excel(roll_single,sheet_name="cluster",startrow=1*len(final_Roll.columns)+3,startcol=1)
	roll_30_60.to_excel(roll_single,sheet_name="cluster",startrow=2*len(final_Roll.columns)+4,startcol=1)
	roll_60_90.to_excel(roll_single,sheet_name="cluster",startrow=3*len(final_Roll.columns)+5,startcol=1)
	roll_90_120.to_excel(roll_single,sheet_name="cluster",startrow=4*len(final_Roll.columns)+6,startcol=1)
	roll_120.to_excel(roll_single,sheet_name="cluster",startrow=5*len(final_Roll.columns)+7,startcol=1)

	row_jump = row_jump + 1

exit()


