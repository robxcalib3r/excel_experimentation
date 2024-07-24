from openpyxl import load_workbook
import pandas as pd

def save2csv(data_arr, file_path, write_column):
    df = pd.read_excel(file_path)
    print(f'len of df: {len(df)}; len of data_arr: {len(data_arr)}')
    for data_idx in range(len(data_arr)):
        df.iloc[data_idx, write_column] = data_arr[data_idx]
    df.to_csv(f'{file_path.split(".")[0]}_mod.csv')

def ID_separate_name(file_path, read_column):
    """
    To separate IDs from the name column
    """
    df = pd.read_excel(file_path)

    ids = []
    for index in range(len(df)):
        givenName = str(df.iloc[index, read_column])
        
        id = ''
        
        for idx in range(len(givenName)-1):
            if givenName[idx].isnumeric() and (givenName[idx+1].isalpha() or ' ' in givenName[idx+1]):
                id = givenName[:idx+1]
                ids.append(id)
            elif givenName.isnumeric():
                id = givenName
                ids.append(id)
                break
        if id == '':
            ids.append(id)

    # print(ids)
    return ids

def ID_separate_accnt(file_path, read_column):
    """
    To separate ID from SamAccountName
    """
    df = pd.read_excel(file_path)

    for index in range(1, len(df)):
        givenName = str(df.iloc[index, read_column])
        
        id = ''
        ids = []
        for idx in range(len(givenName)-1):
            if ('-' in givenName[idx]):
                id = givenName[idx-1:]
                print(id)
                ids.append(id)
    return ids

def cmp_ids(ids_to_cmp, ids_from_list):
    """
    To compare ids from resigned employees (ids_to_cmp) with
    ids from the list of disabled users from AD (ids_from_list)
    """
    ad_status = []
    for resigned_id in ids_to_cmp:
        for idx in range(len(ids_from_list)):
            if str(resigned_id) == str(ids_from_list[idx]):
                ad_status.append("found")
                break
            elif idx == len(ids_from_list)-1:
                ad_status.append('')
    
    return ad_status
        


if __name__ == '__main__':
    disable_path = "disable user list.xlsx"
    resign_path = "Resign_employee_jul23_to_Jun24_report.xls"
    # ids = ID_separate_accnt(path, 4)
    # print(ids)

    ids = ID_separate_name(disable_path, 1)
    # save2csv(ids, disable_path, 2)

    df_resigned = pd.read_excel(resign_path)
    resigned_ids = df_resigned['EMP_ID'].tolist()
    print(len(resigned_ids))

    ad_status = cmp_ids(resigned_ids, ids)

    save2csv(ad_status, resign_path, 10)

    