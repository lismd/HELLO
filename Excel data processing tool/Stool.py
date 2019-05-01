import os
import time
import xlrd
import xlwt
import pandas as pd
import numpy as np

def get_input_file_names(folder_name):
    '''Get Excel file names from a folder'''
    file_list=[]
    for root,dir_names,file_names in os.walk(folder_name):
        for file_name in file_names:
            full_name = os.path.join(root, file_name)
            if full_name.split('.')[-1]=='xls' or full_name.split('.')[-1]=='XLS' or \
            full_name.split('.')[-1]=='xlsx'or full_name.split('.')[-1]=='XLSX':
                file_list.append(full_name)			

    return file_list

def count_sheet_in_excel(excel_file):
    '''Count number of sheets in an excel file'''
    workbook = xlrd.open_workbook(excel_file)
    count = len(workbook.sheets())
    return count
    
def excel_table_by_index(excel_file,sheet_index = 0,colname_row_index = 0,skip_rows = None, skip_footer = 0):
    '''Get data by input sheet index'''
    df = pd.DataFrame(pd.read_excel(excel_file,sheet_name = sheet_index,header = colname_row_index, \
         skiprows=skip_rows,skipfooter = skip_footer)) 
    return df
    
def select_by_columns_index(dataframe,columns_index_use):
    '''Select the data from dataframe by using columns_index_use'''
    columns_name = dataframe.columns.values.tolist()
    columns_use = []
    for index in columns_index_use:
        columns_use.append(columns_name[index])
    dataframe_use = dataframe[columns_use]
    return dataframe_use

def change_column_name(dataframe,new_columns_name):
    '''Columns name of dataframe changed to new_columns_name'''
    row_index_list = list(range(dataframe.shape[0]))
    array_df = np.array(dataframe)
    dataframe_new = pd.DataFrame(array_df,index=row_index_list,columns=new_columns_name)
    return dataframe_new

def add_column_hour_charge(dataframe):
    '''Insert a Column which can show the hours in charge'''
    dataframe.insert(5, 'Charged Hrs in Project', dataframe['Hours'])
    # When Sub Phase equals to 'ONSITE_HOLIDAY', it should be counted into changed hour
    #dataframe['Charged Hrs in Project'].loc[dataframe['Sub Phase'] == 'ONSITE_HOLIDAY'] = 0
    dataframe['Charged Hrs in Project'].loc[dataframe['L1 WBS Des'] == 'ADMIN_AERO'] = 0
    dataframe.insert(7, 'Onsite Holiday', dataframe['Hours'])
    dataframe['Onsite Holiday'].loc[dataframe['Sub Phase'] != 'ONSITE_HOLIDAY'] = 0
    return dataframe
    
def calculate_hours(dataframe):
    '''Calculate total charged hours and total hours of each person'''
    column_name_holiday = ['Onsite Holiday']
    dataframe_holiday = dataframe.groupby(['Cost Ctr','Emp ID'])[column_name_holiday].sum()
    dataframe_holiday_reset = dataframe_holiday.reset_index() 
    dataframe_holiday_sorted = dataframe_holiday_reset.sort_values(['Cost Ctr','Emp ID'])
    column_name = ['Hours','Charged Hrs in Project']
    dataframe_hours = dataframe.groupby(['Cost Ctr','Emp ID'])[column_name].sum()
    dataframe_hours_reset = dataframe_hours.reset_index() 
    dataframe_hours_sorted = dataframe_hours_reset.sort_values(['Cost Ctr','Emp ID'])
    dataframe_hours_sorted.insert(4, 'Onsite Holiday', dataframe_holiday_sorted['Onsite Holiday'])
    return dataframe_hours_sorted

def get_dataframe_with_hours(dataframe,dataframe_hours_new):
    '''Get the dataframe which contains the calculated hours of each person'''
    data_no_duplicate = dataframe.drop_duplicates(subset=['Emp ID'],keep='first',inplace=False).reset_index(drop = True)
    data_no_duplicate_sorted = data_no_duplicate.sort_values(['Cost Ctr','Emp ID']).reset_index()
    data_no_duplicate_sorted.drop(labels=['index'], axis=1,inplace = True)
    data_no_duplicate_sorted.drop(labels=['Hours'], axis=1,inplace = True)
    data_no_duplicate_sorted.insert(4, 'Hours', dataframe_hours_new['Hours'])
    data_no_duplicate_sorted.drop(labels=['Charged Hrs in Project'], axis=1,inplace = True)
    data_no_duplicate_sorted.insert(5, 'Charged Hrs in Project', dataframe_hours_new['Charged Hrs in Project'])
    data_no_duplicate_sorted.drop(labels=['Onsite Holiday'], axis=1,inplace = True)
    data_no_duplicate_sorted.insert(7, 'Onsite Holiday', dataframe_hours_new['Onsite Holiday'])    
    #data_no_duplicate_sorted = data_no_duplicate_sorted.astype({'Hours':'int','Charged Hrs in Project':'int'})
    return data_no_duplicate_sorted

def get_weekly_dataframe(excel_file_list,columns_index_use,new_columns_name):
    '''Get All Excel datas of a folder'''
    dataframe_list=[]
    for excel_file in excel_file_list:
        dataframe = excel_table_by_index(excel_file,0,4,[5])
        dataframe_of_sel_index = select_by_columns_index(dataframe,columns_index_use)
        dataframe_new_column = change_column_name(dataframe_of_sel_index,new_columns_name)
        dataframe_list.append(dataframe_new_column)
    if len(dataframe_list) > 1:
        dataframe_merged = dataframe_list[0].append(dataframe_list[1:],ignore_index = True)
    elif (len(dataframe_list)) == 1:
        dataframe_merged = dataframe_list[0]
    else:
        dataframe_merged = None
    return dataframe_merged

def input_processed_weekly_dataframe(excel_file_list):
    columns_index_use = [4,6,7,18,20,24,11] # index of used colunms name
    columns_use_standard = ['Cost Ctr','Emp ID','Emp Name','Sub Phase','Hours','Target Hrs','L1 WBS Des']
    dataframe_merged = get_weekly_dataframe(excel_file_list,columns_index_use,columns_use_standard)
    dataframe_with_charge = add_column_hour_charge(dataframe_merged)
    dataframe_hours_new = calculate_hours(dataframe_with_charge)
    dataframe_process = get_dataframe_with_hours(dataframe_with_charge,dataframe_hours_new)
    return dataframe_process

def get_weekly_config_dataframe():
    config_file = 'EmployeeInfo.xls'
    dataframe_config = excel_table_by_index(config_file)
    return dataframe_config

def emp_ids_lists(dataframe_weekly_input,dataframe_config):
    input_emp_id_list = list(dataframe_weekly_input['Emp ID'])
    config_emp_id_list = list(dataframe_config['Emp ID'])
    emp_ids_in_config = list(set(input_emp_id_list).intersection(set(config_emp_id_list)))
    emp_ids_not_in_config = list(set(input_emp_id_list).difference(set(config_emp_id_list)))
    return emp_ids_in_config,emp_ids_not_in_config

def save_emp_not_in_config(path ,dataframe):
    #dataframe = pd.DataFrame({'Emp ID':emp_ids_not_in_config})
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    dataframe.to_excel(writer, sheet_name='Sheet1')
    writer.save()

def sort_dataframe_by_colomum(dataframe,columns_name):
    dataframe_sorted = dataframe.sort_values(columns_name).reset_index()
    dataframe_sorted.drop(labels=['index'], axis=1,inplace = True)
    return dataframe_sorted

def combine_input_with_config(dataframe_weekly_input,dataframe_config,emp_ids_in_config):
    same_dataframe_in_input = dataframe_weekly_input[dataframe_weekly_input['Emp ID'].isin(emp_ids_in_config)]
    same_dataframe_in_input = same_dataframe_in_input.reset_index()
    same_dataframe_in_input.drop(labels=['index'], axis=1,inplace = True)

    diff_dataframe_in_input = dataframe_weekly_input[~dataframe_weekly_input['Emp ID'].isin(emp_ids_in_config)]
    diff_dataframe_in_input = diff_dataframe_in_input.reset_index()
    diff_dataframe_in_input.drop(labels=['index'], axis=1,inplace = True)
    diff_dataframe_in_input = diff_dataframe_in_input[['Emp ID','Emp Name']]

    same_dataframe_in_config = dataframe_config[dataframe_config['Emp ID'].isin(emp_ids_in_config)]
    same_dataframe_in_config = same_dataframe_in_config.reset_index()
    same_dataframe_in_config.drop(labels=['index'], axis=1,inplace = True)

    same_dataframe_in_input = sort_dataframe_by_colomum(same_dataframe_in_input,['Emp ID'])
    same_dataframe_in_config = sort_dataframe_by_colomum(same_dataframe_in_config,['Emp ID'])
    
    same_dataframe_in_config.drop(labels=['Target Hrs'], axis=1,inplace = True)
    same_dataframe_in_config.insert(3, 'Target Hrs', same_dataframe_in_input['Target Hrs'])
    same_dataframe_in_config.drop(labels=['Charged Hrs in Project'], axis=1,inplace = True)
    same_dataframe_in_config.insert(5, 'Charged Hrs in Project', same_dataframe_in_input['Charged Hrs in Project'])
    same_dataframe_in_config.drop(labels=['Total Hrs'], axis=1,inplace = True)
    same_dataframe_in_config.insert(6, 'Total Hrs', same_dataframe_in_input['Hours'])
    same_dataframe_in_config.drop(labels=['Onsite Holiday'], axis=1,inplace = True)
    same_dataframe_in_config.insert(12, 'Onsite Holiday', same_dataframe_in_input['Onsite Holiday'])
    dataframe_combine = sort_dataframe_by_colomum(same_dataframe_in_config,['Location','Emp ID'])
    
    return dataframe_combine,diff_dataframe_in_input

def add_total_row(dataframe):
    dataframe.loc[dataframe.shape[0]] = dataframe.apply(lambda x: x.sum())
    index = dataframe.shape[0] - 1
    dataframe.loc[index,'Location'] = 'Total'
    dataframe.loc[index,'Emp ID'] = ''
    dataframe.loc[index,'Employee Name'] = ''
    dataframe.loc[index,'Supervisor'] = ''
    return dataframe
    
def engineer_yield(dataframe):    
    dataframe.drop(labels=['Engineer Yield'], axis=1,inplace = True)
    columns_name = dataframe.columns.values.tolist()
    str_columns_sick_leave = columns_name[8]
    dataframe['Engineer Yield'] = dataframe['Target Hrs'] + dataframe['Overtime Hrs'] + dataframe[str_columns_sick_leave]   \
                                  - (dataframe['Onsite Holiday'])
    dataframe['Engineer Yield'].loc[dataframe['Engineer Yield'] != 0] =              \
                                            1.0 *dataframe['Charged Hrs in Project'] /dataframe['Engineer Yield']
    return dataframe

def finance_yield(dataframe):    
    dataframe.drop(labels=['Finance Yield'], axis=1,inplace = True)
    columns_name = dataframe.columns.values.tolist()
    str_columns_sick_leave = columns_name[8]
    str_columns_vacation_hrs = columns_name[7]
    str_columns_pronl_sick = columns_name[10]
    dataframe['Finance Yield'] = dataframe['Target Hrs'] + dataframe[str_columns_vacation_hrs] +    \
                                 dataframe[str_columns_sick_leave]  + dataframe[str_columns_pronl_sick] 
    dataframe['Finance Yield'].loc[dataframe['Finance Yield'] != 0] =              \
                                            1.0 *dataframe['Charged Hrs in Project'] /dataframe['Finance Yield']
    return dataframe 

def statistic_step(dataframe): 
    dataframe_with_total = add_total_row(dataframe)
    dataframe_EY = engineer_yield(dataframe_with_total)
    dataframe_EY_FY = finance_yield(dataframe_EY)
    return dataframe_EY_FY

def save_dataframe_to_excel(output_path,dataframe,mode_week_month):
    if mode_week_month == "week":
        string_first_row = 'Weekly Engineer Yield Data Statistics'
    else:
        string_first_row = 'Monthly Engineer Yield Data Statistics'
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    dataframe.to_excel(writer, sheet_name='Sheet1',index=False,startrow=1)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Add a header format.
    header_format = workbook.add_format({'text_wrap': True})
    first_row_format = workbook.add_format({
        'bold': True,
        'align':'center',
        'bg_color':'#C8C8C8',
        'font_size':15})

    # Write the column headers with the defined format.
    for col_num, value in enumerate(dataframe.columns.values):
        worksheet.write(1, col_num, value, header_format)

    worksheet.set_row(0, 22) 
    worksheet.merge_range(0,0,0,15,string_first_row,first_row_format)

    worksheet.set_row(1, 120, header_format)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('M:M', 15)
    col_format = workbook.add_format({'num_format': '0.00%'})
    worksheet.set_column('O:O', 10,col_format)
    worksheet.set_column('P:P', 10,col_format)

    writer.save()

def weekly_data_statistic():
    weekly_input_dir = ".\inputData\weekly"
    excel_file_list = get_input_file_names(weekly_input_dir)
    output_dir = ".\\outputData\\weekly\\"

    if len(excel_file_list) > 0:
        dataframe_weekly_input = input_processed_weekly_dataframe(excel_file_list)
        dataframe_config = get_weekly_config_dataframe()
        emp_ids_in_config,emp_ids_not_in_config = emp_ids_lists(dataframe_weekly_input,dataframe_config)
        same_dataframe_in_config,diff_dataframe_in_input = combine_input_with_config(dataframe_weekly_input,dataframe_config,emp_ids_in_config)
        dataframe_EY_FY = statistic_step(same_dataframe_in_config)     
        date_str = time.strftime('%Y-%m-%d',time.localtime(time.time()))    
        output_path = output_dir + date_str + '.xlsx'
        save_dataframe_to_excel(output_path,dataframe_EY_FY,"week")
        if len(emp_ids_not_in_config) > 0:
            ids_not_output_path = output_dir + "Should add below Employee Eids in config file" + '.xlsx'
            save_emp_not_in_config(ids_not_output_path ,diff_dataframe_in_input)
    else:
        print("There is no input data!")
    
def read_monthly_data(excel_file):  
    sheets_count = count_sheet_in_excel(excel_file)
    dataframe_list = []
    for index in range(sheets_count):
        dataframe = excel_table_by_index(excel_file,sheet_index = index,colname_row_index = 1,skip_footer = 1)
        dataframe_list.append(dataframe)
    if len(dataframe_list) > 1:
        dataframe_merged = dataframe_list[0].append(dataframe_list[1:],ignore_index = True)
    else:
        dataframe_merged = dataframe_list[0]
    return dataframe_merged

def get_monthly_dataframe(dataframe_merged):
    key_names = ['Location','Emp ID']
    dataframe_temp = dataframe_merged.drop_duplicates(subset=key_names,keep='first',inplace=False).reset_index(drop = True)
    dataframe_temp = dataframe_temp.sort_values(key_names).reset_index(drop = True)
    dataframe_monthly = dataframe_merged.groupby(key_names).sum()
    dataframe_monthly = dataframe_monthly.sort_values(key_names).reset_index(drop = True)

    dataframe_monthly.insert(0, 'Employee Name', dataframe_temp['Employee Name'])
    dataframe_monthly.insert(0, 'Emp ID', dataframe_temp['Emp ID'])
    dataframe_monthly.insert(0, 'Location', dataframe_temp['Location'])
    dataframe_monthly.insert(12, 'Supervisor', dataframe_temp['Supervisor'])
    dataframe_monthly['Engineer Yield'] = dataframe_temp['Engineer Yield']
    dataframe_monthly['Finance Yield'] = dataframe_temp['Finance Yield']
    return dataframe_monthly

def monthly_data_statistic():
    input_dir = ".\inputData\monthly"
    output_dir = ".\\outputData\\monthly\\"

    date_str = time.strftime('%Y-%m',time.localtime(time.time()))
    output_path = output_dir + date_str + '.xlsx'
    excel_file_list = get_input_file_names(input_dir)
    if len(excel_file_list) > 0:
        excel_file = excel_file_list[0]
        dataframe_merged = read_monthly_data(excel_file)
        dataframe_monthly = get_monthly_dataframe(dataframe_merged)
        dataframe_monthly_final = statistic_step(dataframe_monthly)
        save_dataframe_to_excel(output_path,dataframe_monthly_final,"month")


#######################################################################################################################
# MAIN functions
#######################################################################################################################
weekly_data_statistic()
monthly_data_statistic()
