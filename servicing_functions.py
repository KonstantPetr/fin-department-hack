import json


def get_settings():
    try:
        with open('settings.json', encoding='utf8') as f:
            setting = json.load(f)
        return setting
    except Exception as e:
        return [0, (f'Something wrong with settings file!\n'
                    f'Exception raised: {e}')]


def style_columns(writer, max_col):
    writer.sheets['Изм месяца учёта оказания услуг'].set_column(0, 0, 40)
    writer.sheets['Изм месяца учёта оказания услуг'].set_column(1, 1, 32)
    writer.sheets['Изм месяца учёта оказания услуг'].set_column(2, max_col, 16)
    writer.sheets['Изм даты учёта оказания услуг'].set_column(0, 0, 40)
    writer.sheets['Изм даты учёта оказания услуг'].set_column(1, 1, 32)
    writer.sheets['Изм даты учёта оказания услуг'].set_column(2, max_col, 16)
