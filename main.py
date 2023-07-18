import os
import pandas as pd
import PySimpleGUI as sg
from tqdm import tqdm


bias = ord('A') - 1


def get_excel_column_number(column_name: str) -> int:
    """
    Convert excel column name to column number
    :param column_name: column name
    :return: number of column starting from 1
    """

    column_number = 0
    for i in range(len(column_name)):
        column_number += (ord(column_name[i]) - bias) * (26 ** (len(column_name) - i - 1))
    return column_number


def get_excel_cell_value(df: pd.DataFrame, row: int, column: str) -> str:
    """
    Get excel value by row and column
    :param df: pandas dataframe
    :param row: row number starting from 1
    :param column: column string starting from A
    :return: value of cell
    """
    # convert to uppercase and remove spaces
    column = column.upper().strip()
    return df.iloc[row - 1, get_excel_column_number(column) - 1]

def get_all_sheets_names(excel_path: str) -> [str]:
    """
    Get all sheets names from Excel file
    :param excel_path: path to Excel file
    :return: list of sheets names
    """
    excel_file = pd.ExcelFile(excel_path)
    return excel_file.sheet_names


def write_column_to_output(values: [str], headers_name: [str], output_df: pd.DataFrame) -> pd.DataFrame:
    """
    Write columns to output file
    :param values: list of values
    :param headers_name: list of headers name
    :param output_df: output dataframe
    """
    # append row to output dataframe with value in header
    new_row = pd.DataFrame([values], columns=headers_name)
    output_df = output_df._append(new_row, ignore_index=True)
    return output_df


def init_headers(header_names: [str]) -> pd.DataFrame:
    """
    Init headers in output file
    :param header_names: list of header names
    :return: output dataframe
    """
    return pd.DataFrame(columns=header_names)


def read_config_file(config_file) -> pd.DataFrame:
    """
    Read config file
    :return: config dataframe
    """
    return pd.read_csv(config_file, sep=',')


class GUI:

    def __init__(self):
        sg.theme('DarkAmber')  # Add a touch of color

        # All the stuff inside your window.
        # init window
        self.layout = [
            [sg.Text('选择输入文件:')],
            [sg.Input(key='-INPUT-'), sg.FileBrowse(
                file_types=(("Excel files", "*.xlsx"),), initial_folder=os.getcwd())],
            [sg.Text('选择输出文件夹:')],
            [sg.Input(key='-OUTPUT-DIR-'), sg.FolderBrowse(
                initial_folder=os.getcwd())],
            [sg.Text('选择输出文件名:')],
            [sg.Input(key='-OUTPUT-', default_text='引入模版最新.xlsx')],
            [sg.Text('选择配置文件:')],
            [sg.Input(key='-CONFIG-'), sg.FileBrowse(
                file_types=(("CSV files", "*.csv"),), initial_folder=os.getcwd())],
            [sg.Text('选择Sheet名:')],
            [sg.Input(key='-SHEET-', default_text='销售规格书#单据头(FBillHead)')],
            [sg.Button('开始'), sg.Button('取消')]
        ]

        self.window = sg.Window('Excel to Excel', self.layout, finalize=True)
        self.start_window()


    def _is_valid_path(self, path, file_type):
        if not path:
            return False
        if not os.path.exists(path):
            return False
        if not os.path.isfile(path):
            return False
        if not path.endswith(file_type):
            return False
        return True
    @property
    def input_file(self):
        if not self._is_valid_path(self.window['-INPUT-'].get(), '.xlsx'):
            raise ValueError('Invalid input file')
        return self.window['-INPUT-'].get()

    @property
    def output_file(self):
        output_file = os.path.join(self.window['-OUTPUT-DIR-'].get(), self.window['-OUTPUT-'].get())
        if os.path.exists(output_file):
            raise ValueError('Output file already exists')
        else:
            # create empty file
            open(output_file, 'w').close()
        if not self._is_valid_path(output_file, '.xlsx'):
            raise ValueError('Invalid output file')
        return self.window['-OUTPUT-'].get()

    @property
    def config_file(self):
        if not self._is_valid_path(self.window['-CONFIG-'].get(), '.csv'):
            raise ValueError('Invalid config file')
        return self.window['-CONFIG-'].get()

    @property
    def sheet_name(self):
        return self.window['-SHEET-'].get()


    def start_window(self):
        while True:
            event, values = self.window.read()
            if event == sg.WIN_CLOSED or event == '取消':
                break
            if event == '开始':
                try:
                    start_process(self.input_file, self.output_file, self.config_file, self.sheet_name)
                    sg.popup_ok('处理完成!')
                except Exception as e:
                    sg.popup_error(e)
                    continue
        self.window.close()
        return event, values

def start_process(input_path, output_path, config_file, sheet_name):

    tqdm.write('Reading config file...')
    config_df = read_config_file(config_file)

    headers = config_df['输出列名'].tolist()

    # init output file
    tqdm.write('Init output file...')
    output_df = init_headers(headers)

    sheet_names = get_all_sheets_names(input_path)

    progress_window = sg.Window('进度', [[sg.Text('正在处理...')],
                                         [sg.ProgressBar(len(sheet_names), orientation='h', size=(20, 20), key='progressbar')],])
    progress_window.read(timeout=0)
    progress_bar = progress_window['progressbar']

    for sheet_name in tqdm(sheet_names, desc='Processing sheets'):
        df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
        headers, values = [], []
        for index, row in config_df.iterrows():
            header, row_number, column_header = row[0], row[1], row[2]
            value = get_excel_cell_value(df, row_number, column_header)
            headers.append(header)
            values.append(value)
            output_df = write_column_to_output(values, headers, output_df)
        progress_bar.UpdateBar(sheet_names.index(sheet_name) + 1)
    progress_window.close()

    # write to output file
    tqdm.write('Writing to output file...')

    output_df.to_excel(output_path, index=False, sheet_name=sheet_name)

    tqdm.write('Done!')



if __name__ == '__main__':
    gui = GUI()