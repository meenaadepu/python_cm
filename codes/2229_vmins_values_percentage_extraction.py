import pandas as pd
import os

#df = pd.read_csv(r"D:\Ltv\14LPP\1417\outputs\extracted_files\25C\new_Filtered_FuncMon_Vmin.csv")
df = pd.read_csv(r"D:\Meena\Ltv\22LTV\2229\outputs\extracted_files\andes_vmin.csv")
df_ref_vol = pd.read_excel(r"D:\Meena\Ltv\22LTV\2229\inputs\2214_Andes_yield_vminsearch_Fixed.xlsx", sheet_name='Template')
out_path = r"D:\Meena\Ltv\22LTV\2229\outputs\extracted_files\reports_andes"

# filter_keys ='FM1_7P5T_LVT_CSC28_SAF_VMIN.'
# f_name = filter_keys + "xlsx"
# files_outpath = os.path.join(out_path, f_name)
#
# df_vmin= df.filter(regex=filter_keys, axis=1)
#
# li_vmin=list(df_vmin.columns)
# print(li_vmin)
#
# writer = pd.ExcelWriter(files_outpath, engine='xlsxwriter')

def percent_vmin(df):
    #df = df.iloc[:, 2:]
    f_list = []
    df1_out = pd.DataFrame()
    cols_list = list(df.columns)
    for column in df:
        col_name = column
        print(col_name)
        column = list(df[column])
        voltage = list(df_ref_vol.voltage)
        l = column
        cleanedList = [x for x in column if str(x) != 'nan']
        total_l = len(cleanedList)
        print(total_l)
        if(total_l == 0):
            total_l = 1
        i = j = 0
        for j in range(len(voltage)):
            count = 0
            for i in range(len(l)):
                if (l[i] <= float(voltage[j])):
                    count += 1
            final_values = (count / total_l) * 100
            # print(total_l)
            f_list.append(final_values)
        df_out = pd.DataFrame(f_list)
        f_list = []
        df1_out = pd.concat([df1_out, df_out], axis=1,
                            ignore_index=True, sort=False)
    df1_out.columns = cols_list

    # df1_out = df1_out[['TTm40C','TT125C','FFm40C','FF125C','SSm40C','SS125C','FSm40C','FS125C','SFm40C','SF125C']]
    #
    # cols = pd.MultiIndex.from_arrays(
    #     [['TT', 'TT', 'FF', 'FF', 'SS', 'SS', 'FS', 'FS', 'SF', 'SF'],
    #      ['m40C','125C','m40C','125C','m40C','125C','m40C','125C','m40C','125C']])
    #
    # df1_out.columns = cols

    df1_out = df1_out.round(1)
    df1_out.insert(0, 'Ref_voltage', df_ref_vol["voltage"])
    return df1_out

filter_col_list = ['ANDES_CORE_2_serial_chain_N25_1_pattern_VMIN_SEARCH.','ANDES_CORE_2_serial_chain_N25_2_pattern_VMIN_SEARCH.','ANDES_CORE_2_serial_chain_N25_3_pattern_VMIN_SEARCH.','ANDES_CORE_2_serial_chain_N25_4_pattern_VMIN_SEARCH.','ANDES_CORE_2_serial_chain_N25_5_pattern_VMIN_SEARCH.','ANDES_CORE_3_serial_chain_N25_1_pattern_VMIN_SEARCH.','ANDES_CORE_3_serial_chain_N25_2_pattern_VMIN_SEARCH.','ANDES_CORE_4_serial_chain_N25_1_pattern_VMIN_SEARCH.','ANDES_CORE_5_serial_chain_N25_1_pattern_VMIN_SEARCH.']
# filter_col_list = ['FM_7P5T_LVT_CSC20L_SAF_VMIN.','FM_7P5T_LVT_CSC24L_SAF_VMIN.','FM_7P5T_LVT_CSC28L_SAF_VMIN.','FM_7P5T_LVT_CSC32L_SAF_VMIN.','FM_7P5T_LVT_CSC36L_SAF_VMIN.',
# 'FM_7P5T_SLVT_CSC20SL_SAF_VMIN.','FM_7P5T_SLVT_CSC24SL_SAF_VMIN.','FM_7P5T_SLVT_CSC28SL_SAF_VMIN.','FM_7P5T_SLVT_CSC32SL_SAF_VMIN.','FM_7P5T_SLVT_CSC36SL_SAF_VMIN.',
# 'FM_7P5T_BITCOIN_SLVT_CSC20SL_SAF_VMIN_SEARCH.','FM_7P5T_BITCOIN_SLVT_CSC24SL_SAF_VMIN_SEARCH.','FM_7P5T_BITCOIN_SLVT_CSC28SL_SAF_VMIN_SEARCH.','FM_7P5T_BITCOIN_SLVT_CSC32SL_SAF_VMIN_SEARCH.','FM_7P5T_BITCOIN_SLVT_CSC36SL_SAF_VMIN_SEARCH.']
for filter_col in filter_col_list:
    filter_keys = filter_col
    f_name = filter_keys + "xlsx"
    files_outpath = os.path.join(out_path, f_name)
    df_vmin = df.filter(regex=filter_keys, axis=1)
    li_vmin = list(df_vmin.columns)
    print(li_vmin)
    writer = pd.ExcelWriter(files_outpath, engine='xlsxwriter')

    for i in li_vmin:
       col = str(i)
       print(col)
       df0_index = df.iloc[:,2:8]
       df1 = df[["Split","Temp",col]]
       df1['New'] = df1['Split'].astype(str) + df1['Temp'].astype(str)
       df1=df1.drop(['Split', 'Temp'], axis=1)
       mylist = list(set(df1["New"]))
       df_filter = df_out = pd.DataFrame()
       for j in mylist:
          df2 = df1[df1["New"] == j]
          df_out = pd.concat([df_out, df2.reset_index()[col]], axis=1,
                             ignore_index=True, sort=False)
       print(df_out.shape)
       df_out.columns = mylist

       # df3_out = df_out[['TTm40C','TT125C','FFm40C','FF125C','SSm40C','SS125C','FSm40C','FS125C','SFm40C','SF125C']]
       #
       # col_s = pd.MultiIndex.from_arrays(
       #      [['TT', 'TT', 'FF', 'FF', 'SS', 'SS','FS','FS','SF','SF'],
       #       ['m40C','125C','m40C','125C','m40C', '125C','m40C', '125C', 'm40C', '125C']])
       # df3_out.columns = col_s

       print(col)
       # sht_name = col[0:18]
       sht_name=col.split('_')[-3][3:] #andes



      # sht_name = col.split('_')[-2][3:]  # funcmon

       per_df = percent_vmin(df_out)

       #result = pd.concat([df_out, per_df],axis=1,keys=('Vmins voltage Analysis','Percentage Analysis'),levels=0)

       df_out.to_excel(writer, sheet_name=(sht_name + "_values"))
       per_df.to_excel(writer, sheet_name=sht_name)
       workbook = writer.book
       worksheet = writer.sheets[sht_name]

    writer.save()
    writer.close()
    df_out = pd.DataFrame()

    print('Finished Processing')
