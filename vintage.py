import pandas as pd
import numpy as np

cl_contracts = pd.read_excel("Vintage_input_Common.xlsx",sheet_name="Cl_contracts")
Vintage = pd.read_excel("Vintage_input.xlsx",sheet_name="Count_Vin")
time_iteration = pd.read_excel("Vintage_input_Common.xlsx",sheet_name="VIntage_iteration")
Disbursal = pd.read_excel("combined.xlsx",sheet_name="Disbursal")

Vintage30_amount = pd.read_excel("Vintage_input.xlsx",sheet_name="vin30")
Vintage60_amount = pd.read_excel("Vintage_input.xlsx",sheet_name="vin60")
Vintage90_amount = pd.read_excel("Vintage_input.xlsx",sheet_name="vin90")
Vintage120_amount = pd.read_excel("Vintage_input.xlsx",sheet_name="vin120")

###### Time Mapping ####
time_iteration["Start_date"] = pd.to_datetime(time_iteration["Start_date"])
time_iteration["End_date"] = pd.to_datetime(time_iteration["End_date"])
t_index = len(time_iteration["End_date"])
last_date = time_iteration["End_date"][t_index-1]

## Mapp Contracts + Vintage
def data_preparation(df):
	dummy = pd.merge(left=df,right=cl_contracts,how="left",left_on="CL Contract: CL Contract ID",right_on="CL Contract: CL Contract ID")
	dummy["Contract Date"] = pd.to_datetime(dummy["Contract Date"])
	dummy = dummy.loc[dummy["Contract Date"]<last_date,]
	return dummy

final = data_preparation(Vintage)
final30Amount = data_preparation(Vintage30_amount)
final60Amount = data_preparation(Vintage60_amount)
final90Amount = data_preparation(Vintage90_amount)
final120Amount = data_preparation(Vintage120_amount)

final_master = final.copy()
def DPD_calculation(Dpdlist_1,Dpdlist_2,df):
	df.replace(Dpdlist_1,1,inplace=True)
	df.replace(Dpdlist_2,0,inplace=True)
	return df

final_30_DPD = DPD_calculation(["31-60 DPD","61-90 DPD","91-120 DPD","120+ DPD"],["< 30 DPD","Current"],final)
final=final_master.copy()
final_60_DPD = DPD_calculation(["61-90 DPD","91-120 DPD","120+ DPD"],["31-60 DPD","< 30 DPD","Current"],final)
final=final_master.copy()
final_90_DPD = DPD_calculation(["91-120 DPD","120+ DPD"],["31-60 DPD","61-90 DPD","< 30 DPD","Current"],final)
final=final_master.copy()
final_120_DPD = DPD_calculation(["120+ DPD"],["31-60 DPD","61-90 DPD","91-120 DPD","< 30 DPD","Current"],final)
final=final_master.copy()

def vintage_function_count(input_df):
	final_vinatage =dict()
	total_list = []

	for moblen in range(0,len(time_iteration)):
		final_dummy = input_df.loc[input_df["Contract Date"]>=time_iteration.iloc[moblen,1],]
		final_dummy = final_dummy.loc[final_dummy["Contract Date"]<time_iteration.iloc[moblen,2],]
		total_list.append(len(final_dummy))
		final_dummy.drop(["Contract Date","CL Contract: CL Contract ID"],axis=1,inplace=True)
		final_vinatage[time_iteration.iloc[moblen,0]] = final_dummy.sum()
		final_vinatage[time_iteration.iloc[moblen,0]] = list(reversed(final_vinatage[time_iteration.iloc[moblen,0]]))[moblen:]

	DPD_vintage = pd.DataFrame.from_dict(final_vinatage,orient="index")
	DPD_vintage = DPD_vintage.reset_index()
	DPD_vintage.columns = list(range(0,len(time_iteration["Iteration"])+1))
	DPD_vintage.insert(1,"Total Number of Loans",total_list)
	return DPD_vintage

def vintage_function_percentage(input_df):
	final_vinatage =dict()
	total_list = []

	for moblen in range(0,len(time_iteration)):
		final_dummy = input_df.loc[input_df["Contract Date"]>=time_iteration.iloc[moblen,1],]
		final_dummy = final_dummy.loc[final_dummy["Contract Date"]<time_iteration.iloc[moblen,2],]
		total_list.append(len(final_dummy))
		final_dummy.drop(["Contract Date","CL Contract: CL Contract ID"],axis=1,inplace=True)
		final_vinatage[time_iteration.iloc[moblen,0]] = round(final_dummy.sum()/len(final_dummy)*100,1)
		final_vinatage[time_iteration.iloc[moblen,0]] = list(reversed(final_vinatage[time_iteration.iloc[moblen,0]]))[moblen:]

	DPD_vintage = pd.DataFrame.from_dict(final_vinatage,orient="index")
	DPD_vintage = DPD_vintage.reset_index()
	DPD_vintage.columns = list(range(0,len(time_iteration["Iteration"])+1))
	DPD_vintage.insert(1,"Total Number of Loans",total_list)
	return DPD_vintage

#############################################
########## Amount Calculation ###############
#############################################
def vintage_function_amount(input_df):
	final_vinatage =dict()
	total_list = []

	for moblen in range(0,len(time_iteration)):
		final_dummy = input_df.loc[input_df["Contract Date"]>=time_iteration.iloc[moblen,1],]
		final_dummy = final_dummy.loc[final_dummy["Contract Date"]<time_iteration.iloc[moblen,2],]
		Disbursal_dummy = Disbursal.loc[Disbursal["Contract Date"]>=time_iteration.iloc[moblen,1],]
		Disbursal_dummy = Disbursal_dummy.loc[Disbursal_dummy["Contract Date"]<time_iteration.iloc[moblen,2],]
		total_list.append(sum(Disbursal_dummy["Disbursal Transaction Amount"]))
		final_dummy.drop(["Contract Date","CL Contract: CL Contract ID"],axis=1,inplace=True)
		final_vinatage[time_iteration.iloc[moblen,0]] = final_dummy.sum()
		final_vinatage[time_iteration.iloc[moblen,0]] = list(reversed(final_vinatage[time_iteration.iloc[moblen,0]]))[moblen:]

	DPD_vintage = pd.DataFrame.from_dict(final_vinatage,orient="index")
	DPD_vintage = DPD_vintage.reset_index()
	DPD_vintage.columns = list(range(0,len(time_iteration["Iteration"])+1))
	DPD_vintage.insert(1,"Total Disbursal",total_list)
	return DPD_vintage

def vintage_function_amount_percentage(input_df):
	final_vinatage =dict()
	total_list = []

	for moblen in range(0,len(time_iteration)):
		final_dummy = input_df.loc[input_df["Contract Date"]>=time_iteration.iloc[moblen,1],]
		final_dummy = final_dummy.loc[final_dummy["Contract Date"]<time_iteration.iloc[moblen,2],]
		Disbursal_dummy = Disbursal.loc[Disbursal["Contract Date"]>=time_iteration.iloc[moblen,1],]
		Disbursal_dummy = Disbursal_dummy.loc[Disbursal_dummy["Contract Date"]<time_iteration.iloc[moblen,2],]
		total_list.append(sum(Disbursal_dummy["Disbursal Transaction Amount"]))
		final_dummy.drop(["Contract Date","CL Contract: CL Contract ID"],axis=1,inplace=True)
		final_vinatage[time_iteration.iloc[moblen,0]] = round(final_dummy.sum()/sum(Disbursal_dummy["Disbursal Transaction Amount"])*100,1)
		final_vinatage[time_iteration.iloc[moblen,0]] = list(reversed(final_vinatage[time_iteration.iloc[moblen,0]]))[moblen:]

	DPD_vintage = pd.DataFrame.from_dict(final_vinatage,orient="index")
	DPD_vintage = DPD_vintage.reset_index()
	DPD_vintage.columns = list(range(0,len(time_iteration["Iteration"])+1))
	DPD_vintage.insert(1,"Total Disbursal",total_list)
	return DPD_vintage

writer = pd.ExcelWriter("Vintage.xlsx")
def save_excel(df,name,srow,scol):
	df.to_excel(writer,sheet_name=name,startrow=srow,startcol=scol)
	writer.save()

def main():
	Final30DPD_vintage = vintage_function_count(final_30_DPD)
	save_excel(Final30DPD_vintage,"30DPD",1,1)
	Final30DPD_vintage = vintage_function_percentage(final_30_DPD)
	save_excel(Final30DPD_vintage,"30DPD",len(time_iteration["Iteration"])+3,1)
	Amount30DPD_vintage = vintage_function_amount(final30Amount)
	save_excel(Amount30DPD_vintage,"30DPD",2*len(time_iteration["Iteration"])+5,1)
	Amount30DPD_vintage = vintage_function_amount_percentage(final30Amount)
	save_excel(Amount30DPD_vintage,"30DPD",3*len(time_iteration["Iteration"])+7,1)

	Final60DPD_vintage = vintage_function_count(final_60_DPD)
	save_excel(Final60DPD_vintage,"60DPD",1,1)
	Final60DPD_vintage = vintage_function_percentage(final_60_DPD)
	save_excel(Final60DPD_vintage,"60DPD",len(time_iteration["Iteration"])+3,1)
	Amount60DPD_vintage = vintage_function_amount(final60Amount)
	save_excel(Amount60DPD_vintage,"60DPD",2*len(time_iteration["Iteration"])+5,1)
	Amount60DPD_vintage = vintage_function_amount_percentage(final60Amount)
	save_excel(Amount60DPD_vintage,"60DPD",3*len(time_iteration["Iteration"])+7,1)

	Final90DPD_vintage = vintage_function_count(final_90_DPD)
	save_excel(Final90DPD_vintage,"90DPD",1,1)
	Final90DPD_vintage = vintage_function_percentage(final_90_DPD)
	save_excel(Final90DPD_vintage,"90DPD",len(time_iteration["Iteration"])+3,1)
	Amount90DPD_vintage = vintage_function_amount(final90Amount)
	save_excel(Amount90DPD_vintage,"90DPD",2*len(time_iteration["Iteration"])+5,1)
	Amount90DPD_vintage = vintage_function_amount_percentage(final90Amount)
	save_excel(Amount90DPD_vintage,"90DPD",3*len(time_iteration["Iteration"])+7,1)

	Final120DPD_vintage = vintage_function_count(final_120_DPD)
	save_excel(Final120DPD_vintage,"120DPD",1,1)
	Final120DPD_vintage = vintage_function_percentage(final_120_DPD)
	save_excel(Final120DPD_vintage,"120DPD",len(time_iteration["Iteration"])+3,1)
	Amount120DPD_vintage = vintage_function_amount(final120Amount)
	save_excel(Amount120DPD_vintage,"120DPD",2*len(time_iteration["Iteration"])+5,1)
	Amount120DPD_vintage = vintage_function_amount_percentage(final120Amount)
	save_excel(Amount120DPD_vintage,"120DPD",3*len(time_iteration["Iteration"])+7,1)

main()
# final_60_DPD.to_csv("final.csv",index=False)
