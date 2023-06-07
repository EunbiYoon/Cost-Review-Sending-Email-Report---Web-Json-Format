import pandas as pd
import openpyxl 

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0609/Cost Review_0609.xlsx", sheet_name="TL_PAC")
index_list=list(read_excel["Tool"])
extract_data=read_excel.drop(["Model","Tool"],axis=1)
extract_data.index=range(1,len(extract_data.index)+1)

#column_list
column_list=list(extract_data.columns)

#print start
print("")

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','"nan"')
print(column_json+","+"")

#index_list
row_json=str({"index":index_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','"nan"')
print(row_json+","+"")

#value_list
for i in range(len(extract_data.index)):
    if i==len(extract_data.index)-1:
        value_index="value"+str(i)
        value_list=str(list(extract_data.iloc[i].round(1)))
        value_merge='"'+value_index+'":'+value_list
    else:
        value_index="value"+str(i)
        value_list=str(list(extract_data.iloc[i].round(1)))
        value_merge='"'+value_index+'":'+value_list+","
    print(value_merge)

print("")



