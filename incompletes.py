import pandas as pd
import re
import openpyxl


df=pd.read_csv("Jan_2018.csv")

###############################
#  Data Cleaning Part.
###############################

df["Borrower Industry"] = df["Borrower Industry"].fillna("Not Filled Details")


df["Residence City"] = df["Residence City"].fillna("Not Filled Details")
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Bangalire","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Bangalore","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Bangalore Hoskote","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Bangalore Rural","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Bangalore south","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Bangaluru","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Banglore","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Jayanagar bangalore","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("BENGALORE","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Bengalure","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Bengaluru","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("BOMMASANDRA BENGALURU","Bangalore"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Anjaiah Nagar,Hyderabad","Hyderabad"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Hydearabad","Hyderabad"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Hyderabad","Hyderabad"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Hyderabad East","Hyderabad"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Hyderabad Local","Hyderabad"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Hyderbad","Hyderabad"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Hydrabad","Hyderabad"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Hydrebad","Hyderabad"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Kamothe  Navi Mumbai","Mumbai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Mumbai","Mumbai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Navi Mumbai","Mumbai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("NaviMumbai","Mumbai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Panvel Navi Mumbai","Mumbai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Virar mumbai","Mumbai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Khadki pune","Pune"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Pimpri pune","Pune"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Pune","Pune"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Kalyan dist. Thane","Mumbai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Thane","Mumbai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Thane Kalwa","Mumbai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Chennai","Chennai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Ramapuram chennai","Chennai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Kanchipuram","Chennai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Tiruchirappalli","Chennai"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("K V Rangareddi","Ranga Reddy"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("K.V Rangareddy","Ranga Reddy"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("K.V.Rangareddy","Ranga Reddy"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Ranga Reddy","Ranga Reddy"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Rangareddy","Ranga Reddy"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Visakhapatnam","Visakhapatnam"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Visakhapatnam (Urban)","Visakhapatnam"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Kolkata","Kolkata"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("kolkkata","Kolkata"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("North 24 Parganas","Kolkata"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("South 24 Parganas","Kolkata"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Bardhaman","Kolkata"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("BARASAT","Kolkata"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Central Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("East Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Narela, Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("New Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("North Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("North East Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("North West Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("South Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("South West Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("West Delhi","Delhi"))
df["Residence City"] = df["Residence City"].apply(lambda x: str(x).replace("Gurgaon","Gurugram"))
df.loc[df["Residence City"].isin(["Not Filled Details","Bangalore","Hyderabad","Mumbai","Pune","Mysore","Chennai","Ranga Reddy","Visakhapatnam","Kolkata","Delhi","Gurugram"]) == False,"Residence City"] = "Others"

def age_cut(df):
	df["Age_bin"] = pd.cut(df["Age"],[21,26,31,36,41,45,115],right=False,labels=["21-25","26-30","31-35","36-40","41-45","> 45"])
	df["Age_bin"] = df["Age_bin"].astype("str")
	df.loc[df["Age"].isna() == True,"Age_bin"] = "Not Filled Details"
	return df

df = age_cut(df)

def income_cut(df):
	df["Net Income"]= pd.cut(df["Salary Income-current month"],[0,15001,20001,25001,30001,40001,50001,60001,70001,80001,90001,100001,110001,120001,130001,140001,150001,10000000000],
		right=False,labels=["Below 15,000","15,001  - 20,000","20,001 - 25,000","25,001 - 30,000","30,001 - 40,000","40,001 - 50,000","50,001 - 60,000","60,001 - 70,000","70,001 - 80,000","80,001 - 90,000","90,001 - 100,000","100,001 - 110,000","110,001 - 120,000","120,001 - 130,000","130,001 - 140,000","140,001 - 150,000",">150,000"])
	df["Net Income"] = df["Net Income"].astype("str")
	df.loc[df["Salary Income-current month"].isna() == True,"Net Income"] = "Not Filled Details"
	return df

df = income_cut(df)

# Marital Stages.
df.loc[df["Marital status"].isna() == True,"Marital status"] = "Not Filled Details"
df.loc[df["CL Purpose Name"].isna() == True,"CL Purpose Name"] = "Not Filled Details"

# Education Level
df.loc[df["Education Level"].isna() == True,"Education Level"] = "Not Filled Details"

# UTM Source
df["UTM Source"] = df["UTM Source"].str.lower()
df.loc[df["UTM Source"].isna() == True,"UTM Source"] = "Other"
df.loc[df["UTM Campaign"].isna() == True,"UTM Campaign"] = "Other"
df.loc[df["Residential status"].isna() == True,"Residential status"] = "Not Filled Details"
df.loc[df["No of dependents"].isna() == True,"No of dependents"] = "No of dependents"

#########################################################
#########################################################
#########################################################

list_col = ["Customer ID","Application ID","Borrower Industry","Residence City","Campaign source","Age","Salary Income-current month","Marital status","CL Purpose Name","Education Level","Total years_exp","Residential status","No of dependents","Status","CRIF S1 Score","Reject code","UTM Source","UTM Campaign","Debt service ratio","Active","Age_bin","Net Income"]
df_1=df[list_col]
df_2 = df_1.loc[df["Status"].isin(["FORM INCOMPLETE","FORM COMPLETE, DOCUMENTS PENDING","NEW - ENTERED","NEW - RESOLUTION","NEW - IN REVIEW","LISTED IN MARKETPLACE","REJECTED"])==True,]

df_2.loc[df_2["Campaign source"].isna() == True,"Campaign source"] = "Other"

def save_excel(pivote_tab,col_name):
	pivote_tab = pd.DataFrame(pivote_tab.iloc[:,:-1])
	pivote_prt = pivote_tab.iloc[:-1,:].apply(lambda x: x/x.sum()*100,axis=0)
	pivote_prt = pivote_prt.iloc[:,:].apply(lambda x: x.round(2),axis=1)
	pivote_prt = pd.concat([pivote_tab.iloc[-1:,:],pivote_prt])
	pivote_prt.to_excel(writer,col_name,startrow=1,startcol=1)
	writer.sheets[col_name].column_dimensions.width = 400
	writer.save()
	return 0

# Pivote Tables
writer = pd.ExcelWriter('output.xlsx')

def pivot_making(index_col):
	pivote_tab = pd.pivot_table(data=df_2[[index_col,"Status","Application ID"]],index=index_col,columns="Status",values="Application ID",aggfunc="count",margins=True)
	save_excel(pivote_tab,index_col)
	return 0

# pivote_tab = pd.pivot_table(data=df_2[["Residence City","Status","Application ID"]],index="Residence City",columns="Status",values="Application ID",aggfunc="count")
# save_excel(pivote_tab,"Residence City")

pivote_Tot = pd.pivot_table(data=df_2[["Status","Application ID"]],index="Status",values="Application ID",aggfunc="count")
pivote_Tot.to_excel(writer,"Status Summary",startrow=1,startcol=1)


pivote_head = pd.pivot_table(data=df_2[["Status","Application ID"]],columns="Status",values="Application ID",aggfunc="count")
pivote_head = pivote_head.reset_index()
pivote_head.iloc[0,0] = "Total"
pivote_head = pd.DataFrame(pivote_head)

pivot_making("Borrower Industry")
pivot_making("Residence City")
pivot_making("Campaign source")
pivot_making("Age_bin")
pivot_making("Net Income")
pivot_making("Marital status")
pivot_making("CL Purpose Name")
pivot_making("Education Level")
pivot_making("UTM Source")
pivot_making("UTM Campaign")
pivot_making("Residential status")
pivot_making("No of dependents")
pivot_making("Total years_exp")

