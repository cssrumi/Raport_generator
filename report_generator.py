import openpyxl
from os import path
import pandas as pd
import datetime
from testing_decorators import print_generator, new_timer


def create_file(file_name, dir_path='D:\\PythonRaport\\output'):
    """ Function that create .xlsx file """
    file_path = path.join(dir_path, file_name)
    wb = openpyxl.Workbook()
    wb.save(file_path)


def full_path(file_name, dir_path='D:\\PythonRaport\\output'):
    file_path = path.join(dir_path, file_name)
    return file_path


def validate_and_fix_suffix(file_name, suffix) -> str:
    """ Function that validate the suffix in file and return fixed file_name """
    if suffix[0] != '.':
        suffix = '.' + suffix
    file_suffix = ''
    for char in reversed(file_name):
        file_suffix = char + file_suffix
        if char == '.':
            break
    if file_suffix != suffix:
        file_name = file_name + suffix
    return file_name


# @new_timer
# @print_generator
def validate_repetition(input_data):
    """ Generator(f) that validate and yield not repeated data """
    temp = next(input_data.iterrows())
    for _, d in input_data.iterrows():
        # skip entry with empty username
        if pd.isnull(d[0]):
            continue
        # if username is changed yield and change temp
        if temp[0] != d[0]:
            if not pd.isnull(temp[0]):
                yield temp
            temp = d
        # if username is the same
        # and temp_datetime_start == d_datetime_start => set temp
        elif temp[1] == d[1]:
            temp = d
        # if datetime isn't the same => yield and set temp
        else:
            yield temp
            temp = d


# @new_timer
# @print_generator
def validate_repetition_as_list(input_data):
    """ Generator(f) that validate and yield not repeated data as List """
    temp = ['', '', '']
    for _, d in input_data.iterrows():
        # skip entry with empty username
        if pd.isnull(d['username']):
            continue
        # if username is changed yield and change temp
        if temp[0] != d['username']:
            if temp[0] != '':
                yield temp
            temp = d.tolist()
        # if username is the same
        # and temp_datetime_start == d_datetime_start => set temp
        elif temp[1] == d[1]:
            temp = d.tolist()
        # if datetime isn't the same => yield and set temp
        else:
            yield temp
            temp = d.tolist()


def column_setter(lang: str) -> list:
    column_dict = {
        'PL': ['UserName', 'Data rozpoczęcia txt', 'Godzina rozpoczęcia', 'Długość sesji (min)', 'Data Rozpoczęcia',
               'Imię', 'Nazwisko', 'Płeć', 'Status'],
        'EN': ['UserName', 'Start Date txt', 'Start Time', 'Session length (min)', 'Start Date',
               'Firstname', 'Lastname', 'Sex', 'Status']
    }
    return column_dict.get(lang, column_dict.get('EN'))


def process_datetime(row: list, datetime_format='%Y-%m-%dT%H:%M:%S'):
    """ Function that will return date of start as string, start Hour, session length in minutes and date of start"""
    # list[1] - datetime_start
    # list[2] - datetime_finish
    date_format = '%d.%m.%Y'
    dt1 = datetime.datetime.strptime(row[1], datetime_format)
    dt2 = datetime.datetime.strptime(row[2], datetime_format)
    date = dt1.date().__format__(date_format)
    h = dt1.hour
    m = int(round((dt2 - dt1).seconds / 60, 0))
    return [str(date), h, m, date]


def parse_user_info(username, data):
    for _, d in data.iterrows():
        if username.upper() == d[0].upper():
            r_list = [str(l).replace('\xa0', ' ') for l in d.tolist()]
            return r_list[1:]
    return ['', '', '', '']


def parse_user_info_dict(username, user_dict: dict):
    username = username.upper()
    empty_list = ['', '', '', '']
    value = user_dict.get(username, empty_list)
    return value


@new_timer
def create_user_dict(data) -> dict:
    user_dict = {}
    for _, d in data.iterrows():
        d = [str(l).replace('\xa0', ' ') for l in d.tolist()]
        user_dict[d[0].upper()] = d[1:]
    return user_dict


@new_timer
def create_data_frame(raw_data, user_data, lang='PL'):
    column_list = column_setter(lang)
    df = pd.DataFrame(columns=column_list)
    for value in validate_repetition_as_list(raw_data):
        # new_data.loc[len(new_data)] = value
        row = [value[0]] + process_datetime(value) + parse_user_info(value[0], user_data)
        row_s = pd.Series(row, column_list)
        df = df.append([row_s], ignore_index=True)
    return df


@new_timer
def create_data_frame_dict(raw_data, user_dict: dict, lang='PL'):
    column_list = column_setter(lang)
    df = pd.DataFrame(columns=column_list)
    for value in validate_repetition_as_list(raw_data):
        # new_data.loc[len(new_data)] = value
        row = [value[0]] + process_datetime(value) + parse_user_info_dict(value[0], user_dict)
        row_s = pd.Series(row, column_list)
        df = df.append([row_s], ignore_index=True)
    return df


@new_timer
def main():
    """ 1. Set and validate file_name
        2. Read data from CSV
        3. Read data to compare from xlsx
        4. Process data
        5. Save data in new file"""
    # Set file_name, validate and create that file
    file_name = 'test_file'
    file_name = validate_and_fix_suffix(file_name, 'xlsx')
    # create_file(file_name)

    # Read input_file
    input_file = 'D:\\PythonRaport\\input\\test.csv'
    raw_data = pd.read_csv(input_file, sep=';', iterator=False, header=None)
    # Set Columns name
    raw_data.columns = ['username', 'datetime_start', 'datetime_finish']

    # Read user_data
    user_data_file = 'D:\\PythonRaport\\input\\users.xlsx'
    user_data = pd.read_excel(user_data_file)
    # Create user_dict
    user_dict = create_user_dict(user_data)

    # Create new DataFrame with validated data
    # new_data = create_data_frame(raw_data, user_data)
    new_data = create_data_frame_dict(raw_data, user_dict)

    # Rewrite data in file
    writer = pd.ExcelWriter(full_path(file_name), engine='openpyxl')
    new_data.to_excel(writer, sheet_name='Dane', index=False)
    writer.save()
    return None


if __name__ == '__main__':
    main()
