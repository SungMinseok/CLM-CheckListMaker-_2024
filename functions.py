import os
import time
import pandas as pd
# .py 파일을 모두 찾아내는 함수
def get_recent_file_list(directory, ext = '.py'):
    '''
    os.getcwd()

    
    #print(list(file_dict.keys())[0])
    #print(list(file_dict.values())[0])
    '''
    
    current_directory = directory

    def find_py_files(directory):
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.endswith(ext):
                    yield os.path.join(root, file)

    # 파일 중에서 수정 날짜가 최신인 순으로 정렬
    py_files = list(find_py_files(current_directory))
    py_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

    file_dict = {}

    # 최신 수정 날짜를 가진 파일 출력
    for file in py_files:
        modified_time = os.path.getmtime(file)
        modified_date = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(modified_time))
        file_dict[file] = modified_date

    return file_dict


    

def read_patch_notes(file_path, count = 5):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Sort the DataFrame by the 'Date' column in descending order
        df.sort_values(by='Date', inplace=True, ascending=False)
        # Assuming 'isNotice' is a column in your DataFrame
        filtered_df = df[df['isNotice'] == True]

        top_3_updates = filtered_df.head(count)

        result = []

        # Print the updates in the specified format
        for _, row in top_3_updates.iterrows():
            result.append(f"[{row['Date'].strftime('%y-%m-%d')}] {row['Solution']}")
            
        return '\n'.join(result)

    except Exception as e:
        print(f"Error reading the patch notes: {e}")

    return None