import csv
import pandas as pd
from xlwt import Workbook
import glob

def function(ba,sheet_num,wb,input_file):
    data1 = []
    with open(input_file, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            data1.append(row)
    data1[0][0] = 1
    dictionary = {}
    for i in data1:
        key = int(i[0])
        if key not in dictionary:
            dictionary[key] = {i[1]:float(i[2])}
        else:
            dictionary[key][i[1]] = float(i[2])

    carbon_type = ["can-imports","Coal-IGCC","coal-new","CoalOldScr","CoalOldUns","Gas-CC","Gas-CT","lfill-gas","o-g-s"]
    comparing_data = ["can-imports","Coal-IGCC","coal-new","CoalOldScr","CoalOldUns","Gas-CC","Gas-CT","lfill-gas","o-g-s","biopower","battery","distPV","nuclear","pumped-hydro","upv_2","upv_3","upv_4","upv_5","upv_6","upv_7","dupv_2","dupv_3","dupv_4","dupv_5","dupv_6","dupv_7","wind-ofs_1","wind-ofs_2","wind-ofs_3","wind-ofs_4","wind-ofs_7","wind-ofs_8","wind-ofs_9","wind-ofs_10","wind-ofs_11","wind-ofs_12","wind-ofs_13","wind-ons_1","wind-ons_2","wind-ons_3","wind-ons_4","wind-ons_5","wind-ons_6","wind-ons_7","wind-ons_8","wind-ons_9","wind-ons_10","csp1_12","csp2_10","csp2_11","csp2_12","hydND","hydUD","hydUND","hydNPND","hydED","hydEND","geohydro_pflash_1"]
    set_data = {"upv_2":22,"upv_3":23,"upv_4":24,"upv_5":25,"upv_6":26,"upv_7":27,"dupv_2":28,"dupv_3":29,"dupv_4":30,"dupv_5":31,"dupv_6":32,"dupv_7":33,"wind-ofs_1":11,"wind-ofs_2":12,"wind-ofs_3":13,"wind-ofs_4":14,"wind-ofs_7":15,"wind-ofs_8":16,"wind-ofs_9":17,"wind-ofs_10":18,"wind-ofs_11":19,"wind-ofs_12":20,"wind-ofs_13":21,"wind-ons_1":1,"wind-ons_2":2,"wind-ons_3":3,"wind-ons_4":4,"wind-ons_5":5,"wind-ons_6":6,"wind-ons_7":7,"wind-ons_8":8,"wind-ons_9":9,"wind-ons_10":10,"csp1_12":34,"csp2_10":35,"csp2_11":36,"csp2_12":37}
    gen_total = {}
    year_find = input_file.split('/')
    year_find = year_find[-1]
    year = year_find[0:4]
    year_dict = {2022:1/15,2024:2/15,2026:3/15,2028:4/15,2030:5/15,2032:6/15,2034:7/15,2036:8/15,2038:9/15,2040:10/15,2042:11/15,2044:12/15,2046:13/15,2048:14/15,2050:15/15}
    gen_change = {}
    for i in comparing_data:
        if i in dictionary[ba]:
            gen_total[i] = dictionary[ba][i]
        else:
            gen_total[i] = 0

    gen_total_main = 0
    carbon_type_total = 0

    for i in dictionary:
        if i==ba:
            for j in dictionary[i]:
                gen_total_main += dictionary[i][j]
    gen_total_main+=10**(-28)
    for a in carbon_type:
        carbon_type_total+=gen_total[a]

    percent_carbon=carbon_type_total/gen_total_main
    percent_change = percent_carbon*year_dict[int(year)]
    gen_change_total = gen_total_main*percent_change
    non_carbon_total = gen_total_main-carbon_type_total

    for i in gen_total:
        if i in carbon_type:
            gen_change[i]=gen_total[i]*year_dict[int(year)]

    gen_change['biopower'] = 0.4*gen_change_total
    gen_change['battery'] = 0
    gen_change['nuclear'] = 0
    gen_change['pumped-hydro'] = 0
    sum_off = 0
    viability = {}
    viability['battery'] = 0
    viability['distPV'] = 1
    viability['nuclear'] = 0
    viability['pumped-hydro'] = 0
    viability['hydND'] = 0
    viability['hydUD'] = 0
    viability['hydUND'] = 0
    viability['hydNPND'] = 0
    viability['hydED'] = 0
    viability['hydEND'] = 0
    viability['geohydro_pflash_1'] = 0
    viability['biopower'] = 1

    dataframe = pd.read_excel('viability_matrix.xlsx')

    df_dict = dataframe.to_dict()
    matrix = {}

    for i, j in dataframe.iterrows():
        temp_list = j.to_list()
        matrix[temp_list[0]] = temp_list[1:]
    for i in comparing_data:
        if i not in viability and i in set_data:
            viability[i] = matrix[ba][set_data[i]-1]
    for i in viability:
        sum_off+=viability[i]
    sum_off-=1

    gen_change['distPV'] = gen_change_total*0.6/sum_off
    for i in comparing_data:
        if i not in gen_change:
            if viability[i]==1:
                gen_change[i] = gen_change_total*0.6/sum_off
            else:
                gen_change[i] = 0

    non_base_change = gen_change_total-gen_change['biopower']
    new_gen = {}
    for i in comparing_data:
        if i in carbon_type:
            new_gen[i] = gen_total[i]-gen_change[i]
        else:
            new_gen[i] = gen_total[i] + gen_change[i]

    sheet1 = wb.add_sheet('Sheet '+str(sheet_num))
    for i in range(1,50):
        sheet1.write(i, 0, i)

    sheet1.write(0,1,'renewable_gen_type')
    sheet1.write(0, 2, 'viability')
    variable_count = 1
    for i in viability:
        sheet1.write(variable_count,1,i)
        sheet1.write(variable_count,2,viability[i])
        variable_count+=1

    sheet1.write(0,4,'gen_type')
    sheet1.write(0, 5, 'gen_total')
    sheet1.write(0, 6, 'gen_change')
    sheet1.write(0, 7, 'new_gen')

    variable_count = 1
    for i in gen_change:
        sheet1.write(variable_count,4,i)
        sheet1.write(variable_count, 5, gen_total[i])
        sheet1.write(variable_count, 6, gen_change[i])
        sheet1.write(variable_count, 7, new_gen[i])
        variable_count+=1

    sheet1.write(0,8,'BA')
    sheet1.write(1, 8, ba)
    sheet1.write(0, 9, 'gen_total')
    sheet1.write(1, 9, gen_total_main)
    sheet1.write(0, 10, 'carbon_gen')
    sheet1.write(1, 10, carbon_type_total)
    sheet1.write(0,11,'percent_carbon')
    sheet1.write(1, 11, percent_carbon)
    sheet1.write(0,12,'percent_change')
    sheet1.write(1, 12, percent_change)
    sheet1.write(0,13,'gen_change_total')
    sheet1.write(1, 13, gen_change_total)
    sheet1.write(0, 14, 'percent_change_carbon')
    sheet1.write(1, 14, year_dict[int(year)])
    sheet1.write(0, 15, 'non_carbon_total')
    sheet1.write(1, 15, non_carbon_total)
    sheet1.write(1, 16, gen_change_total/non_carbon_total)
    sheet1.write(0,17,'non_base_change')
    sheet1.write(1, 17, non_base_change)

    return new_gen

files_name = glob.glob("Input_Files/*.csv")
print(files_name)
main_dictionary = {}
row_count = 0
for file in files_name:
    print(file)
    get_name = file.split('/')
    get_name = get_name[-1].split('.')
    wb = Workbook()
    for i in range(1,135):
        catch = function(i,i,wb,file)
        main_dictionary[i] = catch
    put_name = "Output_Files/"+get_name[0]+'_output'+'.xls'
    wb.save(put_name)
    wb2 = Workbook()
    sheet1 = wb2.add_sheet('Main Sheet')
    for i in main_dictionary:
        for j in main_dictionary[i]:
            sheet1.write(row_count,0,i)
            sheet1.write(row_count,1,j)
            sheet1.write(row_count,2,main_dictionary[i][j])
            row_count+=1
    put_name = "Output_Files/" + get_name[0] + '_output_main_sheet' + '.xls'
    wb2.save(put_name)


