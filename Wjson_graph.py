import pandas as pd

############# file name #############
today_date="0721"
week_data="23.07 W3"
model="TL_PAC"
#####################################

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+today_date+"/Cost Review_"+today_date+".xlsx", sheet_name=model)
index_list=list(read_excel["Tool"])
extract_data=read_excel.drop(["Model","Tool"],axis=1)
extract_data.index=range(1,len(extract_data.index)+1)
extract_data=extract_data.drop("Unnamed: 0",axis=1)

#column_list 
column_list=list(extract_data.columns)

#print start
print("")

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','"nan"')
print(column_json+","+"")


#depends model add index 5,6
model_condition=model[:2]
if model_condition=="TL":
    index_list.append("")
elif model_condition=="FL":
    index_list.append("")
    index_list.append("")

#index_list
row_json=str({"index":index_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','"nan"')

# change to the value
print(row_json+","+"")

#value_list
for i in range(len(extract_data.index)):
    if model_condition=="DR":
        #comma end X
        if i==len(extract_data.index)-1:
            value_index="value"+str(i)
            value_list=str(list(extract_data.iloc[i].round(1)))
            value_merge='"'+value_index+'":'+value_list
        #comma end O
        else:
            value_index="value"+str(i)
            value_list=str(list(extract_data.iloc[i].round(1)))
            value_merge='"'+value_index+'":'+value_list+","
    else:
        #comma end O
        value_index="value"+str(i)
        value_list=str(list(extract_data.iloc[i].round(1)))
        value_merge='"'+value_index+'":'+value_list+"," 
    print(value_merge)

# add the value5, value6
if model_condition=="TL":
    print('"value6":[]')
elif model_condition=="FL":
    print('"value5":[],')
    print('"value6":[]')

# add space for copy paste
print("")



