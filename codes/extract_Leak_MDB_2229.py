import pandas as pd
import numpy as np
import json
import pymongo
import os
import re

with open(r'D:\Meena\Ltv\22LTV\config\2229\2229_Leak_extract_config.json') as f:
    config = json.load(f)

# databse global definitions
mng_client = pymongo.MongoClient('localhost', 27017)
mng_db = mng_client[config['db_name']]
db_sim = mng_db[config['db_collection']['sim']]

df_out = pd.DataFrame()


def rotate(l, n):
    return l[n:] + l[:n]


def get_die_filter_data(df, config, process):

    prc_config = [elem for elem in config['wafer_multi'] if elem['process'] == process ]
    out = result = pd.DataFrame()

    for data in prc_config:

        if data['die_values']:  # If die_values are not empty
            print('Not Empty')
            for dx, dy in data['die_values']:
                df_filter = df[(df['Scribe'] == data['scribe']) & (df["Die_X"].apply(lambda x: x == dx)) & (df["Die_Y"].apply(lambda x: x == dy))]

                result = pd.concat([result, df_filter], sort=False, ignore_index=True)
                df_filter.dropna()

            out = pd.concat([out, result], sort=False, ignore_index=True)
            result.dropna()

        else:
            out = pd.concat([out, df], sort=False, ignore_index=True)

    return out

def get_rosc_data(df_filter_data, test_config):
    process, temperature, vd, vn, vp, ro_gname, cname, median, freq, delta, tname, track, chname, min_value, max_value, count = (
        [] for i in range(16))

    for test in list(df_filter_data):
        s = test.split('__', 1)[1].replace('p', '.')
        rx = re.compile(r'-?\d+(?:\.\d+)?')
        volt = rx.findall(s)

        vdd = float('0.' + volt[0].split('.')[1])
        vnw = float(volt[1][1:])
        vpw = float(volt[2])
       #print(test)
        ro_gn = "LEAKAGE_RO%s_TOP" % test.split('_')[1].replace('LEAK', '')
       # print(vdd,ro_gn,vnw,vpw)
        # replaced with db filtering
        df_filter = pd.DataFrame(x for x in db_sim.find({"LEKAGE RO NAME": ro_gn, "TEMP(C)": int(test_config['temp']),
                                                         "Process": test_config['process'], "Voltage(V)": vdd,
                                                         "VNW": vnw, "VPW": vpw}))
        m = df_filter_data[test].median() * 1000000

        if df_filter.empty:
            ref = -1
            d = -1
            cname.append("NA")
            track.append("NA")
            chname.append("NA")

        else:
            df_filter.to_csv("df_filter.csv")
            ref = (df_filter['leakage_iddq(uA)'].values[0])*100
            d = (m / ref)
            chn_name = df_filter['Generic_cell_name'].values[0]
            cname.append(chn_name)
            print("chn",chn_name)
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

    col_labels = ['Split', 'Temp', 'VDD', 'VNW', 'VPW', 'Si-Median', 'Sim', 'Delta', 'Leakage_RO_Name', 'Logic_Cell_used',
                  'Track','Channel', 'Count', 'Min', 'Max', ]
    list_cols = [process, temperature, vd, vn, vp, median, freq, delta,ro_gname, cname, track, chname, count, min_value,
                 max_value]
    zipped = list(zip(col_labels, list_cols))
    df_out = dict(zipped)
    df_rosc = pd.DataFrame(df_out)

    return df_rosc

df1 = pd.DataFrame()
for volt_drive in config['ro_leakage'].keys():

    dir_name = config['out_path']
    main_dir = os.path.join(dir_name, 'final_out')

    if not os.path.exists(main_dir):
        os.makedirs(main_dir)

    for i, sub_config in enumerate(config['ro_leakage'][volt_drive]):

        folder = "%s_%s" % (sub_config['process'], str(sub_config['temp']))

        directory = os.path.join(dir_name, folder)

        if not os.path.exists(directory):
            os.makedirs(directory)

        db_cm = mng_db[config['db_collection'][sub_config['process']]]
        df = pd.DataFrame(x for x in db_cm.find({"TempC": sub_config['temp']}))
        print(df.shape)

        df = get_die_filter_data(df, config, sub_config['process'])

        df1 = pd.concat([df1, df], axis=0, sort=False)


       # df_leak_test = df.filter(regex=('._OFF.'), axis=1)
        df_leak_test = df.filter(regex=('.LEAK.*_OFF.'), axis=1)
        print("df_leak_test",df_leak_test.shape)

     #   df_leak_test = df_leak_test.abs()
        df_leak_test.to_csv("df_leak_test.csv")


        num = df_leak_test._get_numeric_data()
        num[num <= 0] = np.nan
        df_leak_test = num

        f_name = ((sub_config['process'] + '_VDD' + str(sub_config['vdd']) + 'V_VNW' + str(sub_config['vnw']) +
                   'V_VPW' + str(sub_config['vpw']) + 'V_' + str(sub_config['temp'])).replace('.', 'p')
                  .replace('-', 'N')) + "C.xlsx"

        path = os.path.join(directory, f_name)
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        print(f_name)

        test_volt = set('0p' + t.split('__', 1)[1].split('p', 1)[1] for t in list(df_leak_test))  # unique voltage
        df_filter = pd.DataFrame()
        for test in test_volt:
            rx = re.compile(r'-?\d+(?:\.\d+)?')
            volt = rx.findall(test.replace('p', '.'))
            vdd = float(volt[0])
            vnw = float(volt[1][1:])
            vpw = float(volt[2])
            print(vdd,vnw,vpw)
            if vdd == sub_config['vdd'] and vnw == sub_config['vnw'] and vpw == sub_config['vpw']:
                df_filter = df_leak_test.filter(regex='.%s*' % (test), axis=1)
                print("df_filter", df_filter)


        # get the simulation value for the "ROSC_delay" test cases
        df_leak_data = get_rosc_data(df_filter, sub_config)


        df_leak_data.to_excel(writer, index=False)
        df_filter = pd.DataFrame()
        df_out = pd.concat([df_out, df_leak_data])
        writer.save()
        writer.close()
        df_leak_data = pd.DataFrame()

        print("Finished processing : %s " % f_name)
        if i ==0:
            f = (('VDD' + str(sub_config['vdd']) + 'v_VNW' + str(sub_config['vnw']) + 'v_VPW' +
              str(sub_config['vpw']) + 'v').replace('.', 'p').replace('-', 'N')) + ".csv"
    f_path = os.path.join(main_dir, f)

    df_out.to_csv(f_path, index=False)
    df_out = pd.DataFrame()
print("Completed")
