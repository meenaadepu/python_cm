import pandas as pd
import glob
from os.path import join
import numpy as np
import os


root_dir = r"D:\Meena\Ltv\22LTV\2229\outputs\extracted_files"
output_dir = r"D:\Meena\Ltv\22LTV\2229\outputs\extracted_files\reports_andes"
os.chdir(root_dir)

name = "andes_reports.xlsx"
path = os.path.join(output_dir, name)
print(path)
writer = pd.ExcelWriter(path, engine='xlsxwriter')

list_files = []
for file in glob.glob("andes*"):
    fname = (os.path.join(root_dir, file))
    list_files.append(fname)
print(list_files)

for f in list_files:
    #print(f)
    df = pd.read_csv(f)
    process = df['Split'].unique()
    temp = df['Temp'].unique()
    df_col_names = [col for col in df if col.startswith("ANDES")]
   # print(df_col_names)
    l2 = []
    l1 = []
    for k, ele in enumerate(df_col_names):

        for i in process:
            for j in temp:
                l = []
                df_filter = df[(df['Split'] == i) & (df['Temp'] == j)]
                Median = df_filter[ele].median()
                l.append(i)
                l.append(j)
                l.append(ele)
                l.append(Median)
                #print(l)

                l1.append(l)
    col = ['Split','Temp','testcase','Median']
    df_out=pd.DataFrame(l1,columns = col)

    tc = df_out['testcase'].unique()
    df3 = df1 = df2 =df_final=df_l=pd.DataFrame()
    l1=[]
    for i in tc:
        df1=df_out[df_out['testcase']==str(i)]
        l1 = list(df1.Median.values)
        df1.rename(columns={'Median': i}, inplace=True)
        df2[i] = l1

    print(df1.head(3))
    print(list(df1.columns))
    #df_l = df1.iloc[:,0:4]
    df_l = df1[["Split","Temp"]]
    print(df_l.head())
    df2.insert(0, "Split", list(df_l["Split"]))
    df2.insert(1, "Temp", list(df_l["Temp"]))

    #f_name = f.split(".")[0] + "_summary" + ".csv"
    # f_name = output_dir+"\\" + f.split("\\")[-1].split(".")[0] + "_summary" + ".csv"
    print(f)
    f_name = f.split("\\")[-1].split(".")[0]
    print(f_name)

    df2.to_excel(writer, sheet_name=f_name, index=False)
    workbook = writer.book
    worksheet = writer.sheets[f_name]

    #df2.to_csv(f_name,index=True)
    print("finished")
    df = pd.DataFrame()
writer.save()
writer.close()
