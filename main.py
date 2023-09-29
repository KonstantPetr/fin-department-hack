from servicing_functions import get_settings
from get_google_sheet import get_google_sheet
from data_manipulation_middleware import get_actual_data, assemble_data


if __name__ == '__main__':

    settings = get_settings()
    if type(settings) is list:
        print(settings)
        exit()

    google_table_check = get_google_sheet(settings['GOOGLE_TABLE_ID'])
    if google_table_check != 'Google table acquired!':
        print(google_table_check)
        exit()

    df = get_actual_data(settings['SHEET_NAME'], settings['COLUMNS'])

    print(assemble_data(df))
