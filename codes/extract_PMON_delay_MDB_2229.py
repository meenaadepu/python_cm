import pandas as pd
import json
import pymongo
import os

with open(r'D:\Meena\Ltv\22LTV\config\2228_Die1\2228die1_PMON_Extract_config_single.json') as f:
    config = json.load(f)

# databse global definitions
mng_client = pymongo.MongoClient('localhost', 27017)
mng_db = mng_client[config['db_name']]
db_sim = mng_db[config['db_collection']['sim']]

def rotate(l, n):
    return l[n:] + l[:n]

def get_die_filter_data(df,config, process):
    prc_config = [elem for elem in config['wafer_multi'] if elem['process'] == process]
    out = result = pd.DataFrame()
    print (prc_config)

    for data in prc_config:
      #  print(df['Scribe'],data['scribe'])
        if not data['die_values']:  # If die_values are empty

            df_filter = df[(df['Scribe'] == data['scribe']) &(df['Wafer'] == data['alias']) &
                           (df['HardBin'] == 1) & (df['SoftBin'] == 1)]
            result = pd.concat([result, df_filter], sort=False, ignore_index=True)
            df_filter =pd.DataFrame()

        else:  # if die_values are available
            print ('Dievalues avail')
            for dx, dy in data['die_values']:

                df_filter = df[(df['Scribe'] == data['scribe']) &(df['Wafer'] == data['alias']) & (df['HardBin'] == 1) & (df['SoftBin'] == 1) &
                               (df["Die_X"].apply(lambda x: x == dx)) & (df["Die_Y"].apply(lambda x: x == dy))]

                result = pd.concat([result, df_filter], sort=False, ignore_index=True)
                df_filter=pd.DataFrame()

        out = pd.concat([out, result], sort=False, ignore_index=True)
        result=pd.DataFrame()

    return out


def get_rosc_data(df_filter_data, test_config):
   # print(df_filter_data.shape)
    process, temperature, vd, vn, vp, ro_gname, cname, median, freq, delta, tname, track, chname, min_value, max_value, count = (
        [] for i in range(16))

    for test in list(df_filter_data):
        volt = test.split('__')[1] + "_" + test.split('__')[2]
        vdd = float(tes.split('__')[1][5:].replace("p", "."))
        vnw = (tes.split('__')[2].split('_')[1][3:].replace("p", "."))
        vpw = (tes.split('__')[2].split('_')[3][3:].replace("p", "."))
        ro_gn = test.split('__')[3].split("[")[0]

        print(vdd,vnw,vpw,ro_gn)
        # replaced with db filtering
        df_filter = pd.DataFrame(x for x in db_sim.find({"RO_cell_name": ro_gn, "TEMP(C)": (test_config['temp']),
                                                         "Process": test_config['process'], "Voltage(V)": vdd,
                                                         "VNW": vnw, "VPW": vpw}))

        m = df_filter_data[test].median()

        if df_filter.empty:
            ref = -1
            d = -1
            cname.append("NA")
            track.append("NA")
            chname.append("NA")

        else:


            ref = float(df_filter['frequency(MHz)'].values[0]) / 32 * 1000000

            d = ((ref - m) / ref)

            chn_name = df_filter['Generic_cell_name'].values[0]
            cname.append(chn_name)
            track.append((chn_name.split("_")[0][2:]).replace("P", "."))
            n = chn_name.split("_")
            chname.append(rotate(n, len(n) - 1)[0])

        process.append(test_config['process'])
        temperature.append(test_config['temp'])
        vd.append(vdd)
        vn.append(vnw)
        vp.append(vpw)
        ro_gname.append(ro_gn)
        median.append(m)
        freq.append(ref)
        delta.append(d)
        tname.append(test)
        min_value.append(df_filter_data[test].min())
        max_value.append(df_filter_data[test].max())
        count.append(df_filter_data[test].count())

    col_labels = ['RO_Generic_Name', 'Logic_Cell_used', 'Track', 'Channel', 'Count', 'Min', 'Max', 'Median',
                  'Sim value', 'Delta', ]
    list_cols = [ro_gname, cname, track, chname, count, min_value, max_value, median, freq, delta]
    zipped = list(zip(col_labels, list_cols))
    df_out = dict(zipped)
    df_rosc = pd.DataFrame(df_out)

    return df_rosc


for v1, v2 in config['bias']:
    print(v1,v2)

    for sub_config in config['rosc_delay_config']:

        dir_name = config['out_path']
        folder = "%s_%s" % (sub_config['process'], str(sub_config['temp']))

        directory = os.path.join(dir_name, folder)

        if not os.path.exists(directory):
            os.makedirs(directory)

        db_cm = mng_db[config['db_collection'][sub_config['process']]]
        df = pd.DataFrame(x for x in db_cm.find({"TempC": sub_config['temp'], "HardBin": 1, "SoftBin": 1}))
        #df = get_die_filter_data(df,config,sub_config['process'])

        # df1 = pd.concat([df1,df],axis=0,sort=False)
        # df1.to_csv(r"D:\Ltv\22LTV\2226\outputs\2226_5per_diefilter.csv")

        rosc_test = df.filter(regex=('PMON_DA_ACESS_M0_W1_IP1.*'), axis=1)
        print(rosc_test.shape)
        test_set = set([x.split(':')[0] for x in list(rosc_test)])
       # print(test_set)
        f_name = ((sub_config['process'] + '_VDD' + str(sub_config['vdd']) + 'V_VNW' + str(v1) + 'V_VPW' +
                   str(v2) + 'V_' + str(sub_config['temp'])).replace('.', 'p').replace('-', 'N')) + "C.xlsx"

        path = os.path.join(directory, f_name)

        writer = pd.ExcelWriter(path, engine='xlsxwriter')


        for test_name in test_set:

            df_rosc_filter = df[df.columns[pd.Series(df.columns).str.startswith(test_name)]]
            test_grp = [y for y in set(list(rosc_test)) if y.split(':')[0] == test_name]  # track all voltage tests
            test_volt = set([z.split('__w1_ip1_da')[0] for z in test_grp])  # unique voltage
            #print("test_volt",test_volt)
            for tes in test_volt:
                df_t_filter =df_rosc_data = pd.DataFrame()
                print(tes)
                vdd = float(tes.split('__')[1][5:].replace("p", "."))
                vnw = (tes.split('__')[2].split('_')[1][3:].replace("p", "."))
                vpw = (tes.split('__')[2].split('_')[3][3:].replace("p", "."))
                print(vdd,vnw,vpw)

                if vdd == sub_config['vdd'] and vnw == float(v1) and vpw == float(v2):
                    print("mmmmmmmmm")

                    df_t_filter = df_rosc_filter.filter(regex=tes + '.*', axis=1)

                    print("df_t_filter",df_t_filter.shape)
                    df_t_filter.to_csv("df_t_filter_rosc.csv")

                   # get the simulation value for the "ROSC_delay" test cases
                    df_rosc_data = get_rosc_data(df_t_filter, sub_config)
                    print("df_rosc_data",df_rosc_data.shape)


                    df_rosc_data.to_excel(writer, sheet_name=test_name, index=False)
                    df_t_filter=pd.DataFrame()
                    df_rosc_filter=pd.DataFrame()

        writer.save()
        writer.close()
        df_rosc_data = pd.DataFrame()

        print ("Finished processing : %s " %f_name)

    print("Completed")
