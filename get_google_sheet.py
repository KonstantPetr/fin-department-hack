import urllib.request


def get_google_sheet(google_table_id):
    destination = 'actual_work_table.xlsx'
    url = f'https://docs.google.com/spreadsheets/d/{google_table_id}/export?exportFormat=xlsx'
    try:
        urllib.request.urlretrieve(url, destination)
    except Exception as e:
        return (f'Something wrong with url!\n'
                f'Exception raised: {e}')
    return 'Google table acquired!'
