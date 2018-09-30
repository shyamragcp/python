import pandas as pd
import numpy as np
import datetime
import calendar
from collections import Counter
from datetime import date
from datetime import datetime
import re
import openpyxl

df = pd.read_csv("main.csv")

# print(len(df["Customer ID"]))

df = df.loc[df.Status.isin(["LISTED IN MARKETPLACE","REJECTED","NEW - ENTERED","NEW - RESOLUTION","APPROVED - APPROVED","NEW - IN REVIEW","NEW - IN CREDIT CHECK"]) == True,]
df = df.loc[df["Application Completion time"].isna() == False,]
# print(len(df["Customer ID"]))
# city= pd.read_csv("city_final.csv")
# df = pd.merge(df,city,how="left",left_on="Customer ID",right_on="Customer ID")


def time_stamp(final,date_column,new_column):
	final[date_column] = pd.to_datetime(final[date_column])
	final[new_column+"Year"]=[x.year for x in final[date_column]]
	final[new_column+"Month"]=[x.month for x in final[date_column]]
	final.sort_values([new_column+"Year",new_column+"Month"],inplace=True)
	final[new_column+"Time"] = [calendar.month_abbr[int(x)] for x in final[new_column+"Month"]]
	final["Time Stamp"] =  final[new_column+"Time"].map(str)+" "+ final[new_column+"Year"].map(str)
	return final



df = time_stamp(df,"Application Completion time",new_column="app_")


# df["CRIF Score"] = df["CRIF Score"].apply(lambda x: str(x).split(",")[0])
# df["CRIF Score"] = df["CRIF Score"].apply(lambda x: str(x).replace("[a-zA-Z]",""))
# df.loc[df["CRIF S1 Score"].isna() == True,"CRIF S1 Score"] = df.loc[df["CRIF S1 Score"].isna() == True,"CRIF Score"]


# Cleaning For City Major Cities.

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


df.loc[df["Residence City"].isin(["Bangalore","Hyderabad","Mumbai","Pune","Mysore","Chennai","Ranga Reddy","Visakhapatnam","Kolkata","Delhi","Gurugram"]) == False,"Residence City"] = "Others"

# print(len(df["Customer ID"]))

pivote = pd.pivot_table(data=df,index=["app_Year","app_Month"],columns="Residence City",aggfunc="count",values="Application ID")
pivote = pd.DataFrame(pivote)

# pivote["col"] = pd.Series(pivote.index[0][1]).map(str) + pd.Series(pivote.index[0][0]).map(str)
# pivote.index[0][1] = calendar.month_abbr[int(pivote.index[0][1])]

# print(pivote.head())
# exit()

df.to_csv("city_test.csv")
exit()

# 
index_list = ["Sep 2016",
"Oct 2016",
"Nov 2016",
"Dec 2016",
"Jan 2017",
"Feb 2017",
"Mar 2017",
"Apr 2017",
"May 2017",
"Jun 2017",
"Jul 2017",
"Aug 2017",
"Sep 2017",
"Oct 2017",
"Nov 2017",
"Dec 2017",
"Jan 2018",
"Feb 2018",
"Mar 2018",
"Apr 2018",
"May 2018",
"Jun 2018",
"Jul 2018",
"Aug 2018",
"Sep 2018"]

pivote["index_col"] = index_list

pivote.reset_index(inplace=True)
pivote.drop(["app_Year","app_Month"],inplace=True,axis=1)
pivote.set_index("index_col",inplace=True)

pvote_1 = pivote.loc[:,:].apply(lambda x: x/x.sum()*100,axis=1)
pvote_1 = pvote_1.loc[:,:].apply(lambda x: x.round(2),axis=1)

writer = pd.ExcelWriter('output.xlsx')
pivote.to_excel(writer,"Total City Frequency Count")
pvote_1.to_excel(writer,"Total City Frequency Percentage")
writer.save()


# print(pvote_1.index)
# exit()

import dash
import dash_core_components as dcc
import dash_html_components as html
import plotly.graph_objs as go


app = dash.Dash()

app.layout = html.Div(children=[

	html.H2(children = "City",style={"textAlign":"left"}),
	html.H3(children = "Base: Total Application",style={"textAlign":"right"}),
	dcc.Graph(
		id = "plot_P1",
		figure = {
		"data" : [go.Bar(y=pvote_1.index,x=pvote_1["Bangalore"],name="Bangalore",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Chennai"],name="Chennai",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Delhi"],name="Delhi",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Gurugram"],name="Gurugram",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Hyderabad"],name="Hyderabad",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Kolkata"],name="Kolkata",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Mumbai"],name="Mumbai",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Mysore"],name="Mysore",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Pune"],name="Pune",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Ranga Reddy"],name="Ranga Reddy",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Visakhapatnam"],name="Visakhapatnam",orientation = "h"),
		go.Bar(y=pvote_1.index,x=pvote_1["Others"],name="Others",orientation = "h")
		],
		"layout":go.Layout(
			xaxis = {"title":"Percentage"},
			yaxis = {"title":"Time"},
			barmode="stack"
			)
		}
		),

	html.H2(children = "City",style={"textAlign":"left"}),
	html.H3(children = "Base: Total Application",style={"textAlign":"right"}),
	dcc.Graph(
		id = "plot_P2",
		figure = {
		"data" : [go.Bar(x=pvote_1.index,y=pvote_1["Bangalore"],name="Bangalore"),
		go.Bar(x=pvote_1.index,y=pvote_1["Chennai"],name="Chennai"),
		go.Bar(x=pvote_1.index,y=pvote_1["Delhi"],name="Delhi"),
		go.Bar(x=pvote_1.index,y=pvote_1["Gurugram"],name="Gurugram"),
		go.Bar(x=pvote_1.index,y=pvote_1["Hyderabad"],name="Hyderabad"),
		go.Bar(x=pvote_1.index,y=pvote_1["Kolkata"],name="Kolkata"),
		go.Bar(x=pvote_1.index,y=pvote_1["Mumbai"],name="Mumbai"),
		go.Bar(x=pvote_1.index,y=pvote_1["Mysore"],name="Mysore"),
		go.Bar(x=pvote_1.index,y=pvote_1["Pune"],name="Pune"),
		go.Bar(x=pvote_1.index,y=pvote_1["Ranga Reddy"],name="Ranga Reddy"),
		go.Bar(x=pvote_1.index,y=pvote_1["Visakhapatnam"],name="Visakhapatnam"),
		go.Bar(x=pvote_1.index,y=pvote_1["Others"],name="Others")
		],
		"layout":go.Layout(
			xaxis = {"title":"Time"},
			yaxis = {"title":"Percentage"},
			barmode="stack"
			)
		}
		)
	])


app.run_server(debug=True)

