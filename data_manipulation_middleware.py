import pandas as pd
from datetime import date
from servicing_functions import style_columns


def get_actual_data(sheet_name, columns):
    df = pd.read_excel('actual_work_table.xlsx', sheet_name=sheet_name, usecols=columns)
    return df


def get_previous_output():
    old_df_1 = pd.read_excel('output_tables.xlsx',
                             sheet_name='Изм месяца учёта оказания услуг').iloc[:, 2:].copy()
    old_df_2 = pd.read_excel('output_tables.xlsx',
                             sheet_name='Изм даты учёта оказания услуг').iloc[:, 2:].copy()
    return [old_df_1, old_df_2]


def assemble_data(input_df):
    pd.set_option('display.max_columns', None)
    today = date.today().strftime('%d.%m.%Y')
    df_1 = input_df.iloc[:, [0, 3]].copy()
    df_2 = input_df.iloc[:, [0, 3]].copy()

    try:
        previous_output = get_previous_output()
        old_df_1, old_df_2 = previous_output[0], previous_output[1]

        for x in range(old_df_1.shape[0], df_1.shape[0]):
            old_df_1.loc[x] = ['' for y in range(old_df_1.shape[1])]
            old_df_2.loc[x] = ['' for y in range(old_df_2.shape[1])]

        df_1 = pd.concat([df_1, old_df_1], sort=False, axis=1)
        df_2 = pd.concat([df_2, old_df_2], sort=False, axis=1)
        today_data_1 = []
        today_data_2 = []

        for x in range(df_1.shape[0]):

            for col_index in range(-1, -old_df_1.shape[1] - 1, -1):
                if not (col_index + 1):
                    if type(old_df_1.iloc[:, col_index:].squeeze()[x]) != str:
                        if col_index == -old_df_1.shape[1]:
                            today_data_1.append(input_df.iloc[:, 2][x])
                        pass
                    elif old_df_1.iloc[:, col_index:].squeeze()[x] == '':
                        today_data_1.append(input_df.iloc[:, 2][x])
                        break
                    elif input_df.iloc[:, 2][x] != old_df_1.iloc[:, col_index:].squeeze()[x]:
                        today_data_1.append(input_df.iloc[:, 2][x])
                        break
                    else:
                        today_data_1.append(float('nan'))
                        break

                else:
                    if type(old_df_1.iloc[:, col_index:col_index + 1].squeeze()[x]) != str:
                        if col_index == -old_df_1.shape[1]:
                            today_data_1.append(input_df.iloc[:, 2][x])
                        pass
                    elif old_df_1.iloc[:, col_index:col_index + 1].squeeze()[x] == '':
                        today_data_1.append(input_df.iloc[:, 2][x])
                        break
                    elif input_df.iloc[:, 2][x] != old_df_1.iloc[:, col_index:col_index + 1].squeeze()[x]:
                        today_data_1.append(input_df.iloc[:, 2][x])
                        break
                    else:
                        today_data_1.append(float('nan'))
                        break

            for col_index in range(-1, -old_df_2.shape[1] - 1, -1):
                if not (col_index + 1):
                    if type(old_df_2.iloc[:, col_index:].squeeze()[x]) != str:
                        if col_index == -old_df_2.shape[1]:
                            today_data_2.append(input_df.iloc[:, 1][x].strftime('%d.%m.%Y'))
                        pass
                    elif old_df_2.iloc[:, col_index:].squeeze()[x] == '':
                        today_data_2.append(input_df.iloc[:, 1][x].strftime('%d.%m.%Y'))
                        break
                    elif input_df.iloc[:, 1][x].strftime('%d.%m.%Y') != \
                            old_df_2.iloc[:, col_index:].squeeze()[x]:
                        today_data_2.append(input_df.iloc[:, 1][x].strftime('%d.%m.%Y'))
                        break
                    else:
                        today_data_2.append(float('nan'))
                        break

                else:
                    if type(old_df_2.iloc[:, col_index:col_index + 1].squeeze()[x]) != str:
                        if col_index == -old_df_2.shape[1]:
                            today_data_2.append(input_df.iloc[:, 1][x].strftime('%d.%m.%Y'))
                        pass
                    elif old_df_2.iloc[:, col_index:col_index + 1].squeeze()[x] == '':
                        today_data_2.append(input_df.iloc[:, 1][x].strftime('%d.%m.%Y'))
                        break
                    elif input_df.iloc[:, 1][x].strftime('%d.%m.%Y') != \
                            old_df_2.iloc[:, col_index:col_index + 1].squeeze()[x]:
                        today_data_2.append(input_df.iloc[:, 1][x].strftime('%d.%m.%Y'))
                        break
                    else:
                        today_data_2.append(float('nan'))
                        break

        df_1[today], df_2[today] = today_data_1, today_data_2

    except FileNotFoundError:
        print('There is no previous data! Launch the initial procedure!')
        df_1[today] = [x for x in input_df.iloc[:, 2]]
        df_2[today] = [x.strftime('%d.%m.%Y') for x in input_df.iloc[:, 1]]

    try:
        writer = pd.ExcelWriter('output_tables.xlsx', engine='xlsxwriter')
        df_1.to_excel(writer, sheet_name='Изм месяца учёта оказания услуг', index=False)
        df_2.to_excel(writer, sheet_name='Изм даты учёта оказания услуг', index=False)
        style_columns(writer, df_1.shape[0] - 1)
        writer.close()
        return 'Success!'
    except PermissionError:
        return 'Table is busy! Close file and try again!'
