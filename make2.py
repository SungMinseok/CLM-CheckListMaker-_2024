import pandas as pd
from openpyxl import load_workbook
import time
from tqdm import tqdm
import os
import re
from xlsx_processing import *
def make_dataframe_to_dict(data_df):
    result_dict = {col: [] for col in data_df.columns}

    for index, row in data_df.iterrows():
        for col_name, col_value in row.items():            
            #print(f'{col_name=}')
            if not pd.isna(col_value):
                result_dict[col_name].append(col_value)

    return result_dict



def process_data_template(data_file, template_file, data_sheet_name, sheet_name, key_column, result_file_name):
    #cur_time = time.strftime('%Y%m%d_%H%M%S')
    #result_file_name = os.path.join(result_path, f"{sheet_name}_{cur_time}.xlsx")
    
    data_df = pd.read_excel(data_file, sheet_name=data_sheet_name)
    #print(data_df.columns)

    targetIdList = data_df[key_column].dropna(axis=0) #실제 ID의 리스트
    targetIdIndexList = targetIdList.index
    totalCount = len(targetIdIndexList)

    template_df = pd.read_excel(template_file, sheet_name=sheet_name)

    final_df = pd.DataFrame()

    for j in tqdm(range(0,totalCount)):
        if (j+1) >= len(targetIdIndexList) :
            tempDf = data_df[targetIdIndexList[j]:]
        else :
            tempDf = data_df[targetIdIndexList[j]:targetIdIndexList[j+1]]
        tempDf = tempDf.reset_index()

        cur_dict = make_dataframe_to_dict(tempDf)
        output_df = template_df.copy(deep=True)
        no_repeat_col_list = []

        for col_name, col_value in cur_dict.items():

            #열 값의 데이터가 1개일 경우
            if len(col_value) == 1 :
                placeholder = f'{{{col_name}}}'

                try:
                    output_df.replace(f'{{{col_name}}}', f'{int(col_value[0])}',regex=True,inplace=True)
                except:
                    output_df.replace(f'{{{col_name}}}', f'{str(col_value[0])}',regex=True,inplace=True)
                    
            #열 값의 데이터가 2개 이상일 경우
            else :


                placeholder = f'{{{col_name}}}'
                
                #if col_name not in no_repeat_col_list :
                if len(no_repeat_col_list)== 0 :
                    target_cell_values = output_df.stack()[output_df.stack().str.contains(placeholder)].tolist()

                for target_cell_value in target_cell_values :
                    new_value_list = []


                    if len(no_repeat_col_list)== 0 :
                        matches = re.findall(r'{(.*?)}', target_cell_value)
                        for match in matches :
                            no_repeat_col_list.append(match)
                        no_repeat_col_list.remove(col_name)


                    for i, value in enumerate(col_value) :#['R2M 1주년 감사패[이벤트]', '다이아 상자']
                        #반복시키는것의 처음.(이름)
                        if col_name not in no_repeat_col_list :
                            placeholder = f'{{{col_name}}}'
                            
                            
                            temp_value = target_cell_value.replace(placeholder,f'{str(value)}')


                            for no_repeat_col_name in no_repeat_col_list:
                                placeholder = f'{{{no_repeat_col_name}}}'
                                temp_value = temp_value.replace(placeholder,f'{{{no_repeat_col_name}_{i}}}')
                            if temp_value not in new_value_list :
                                new_value_list.append(temp_value)

                        #반복시키는것의 2번째 이상.(개수 등)
                        else:
                                                    
                            col_name = f'{col_name}'
                            output_df.replace(placeholder, f'{str(value)}',regex=True,inplace=True)
                            output_df.replace(f'{{{col_name}_{i}}}', f'{str(value)}',regex=True,inplace=True)

                    #target_cell_values.remove(target_cell_value)

                    
                    temp_list_0 = '\n'.join(new_value_list).split('\n')
                    temp_list_1 = []
                    for value in temp_list_0:
                        if value not in temp_list_1:
                            temp_list_1.append(value)
                            no_repeat_col_list.append(col_name)

                    if len(new_value_list) != 0 :
                        output_df.replace(target_cell_value,'\n'.join(temp_list_1), regex=False,inplace=True)
                        #output_df.replace(target_cell_value,'\n'.join(set(temp_list)), regex=False,inplace=True)
                
                    no_repeat_col_list = []
        final_df = pd.concat([final_df,output_df])

    final_df.to_excel(result_file_name, index=False)
    
    postprocess_cashshop(result_file_name)
    #os.startfile(result_file_name)

if __name__ == "__main__" :
    data_file='data.xlsx'
    template_file= 'template.xlsx'

    #data_sheet_name = 'Event'
    #data_sheet_name = 'Cashshop'
    data_sheet_name = '길드도감'
    #sheet_name ='Event'
    #sheet_name ='Cashshop'
    sheet_name ='길드도감재료'

    #sheet_name ='Cashshop (2)'
    #key_column = 'EventName'
    #key_column = 'CashShop ID'
    #key_column = 'Category'
    key_column = '도감 이름'
    #key_column = '길드 레벨'
    result_path = 'test'
    cur_time = time.strftime('%Y%m%d_%H%M%S')
    result_file_name = os.path.join(result_path, f"{sheet_name}_{cur_time}.xlsx")
    process_data_template(data_file, template_file, data_sheet_name, sheet_name, key_column,'test.xlsx')
