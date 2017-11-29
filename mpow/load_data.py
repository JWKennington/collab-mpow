'''
This module contains utilities for loading and parsing data for MPOW. The functions below parse data from Excel sheets (xls format). The 
data are arranged such that each row represents the time-ordered observations for a particular patient. In the pain score data, the 
color green represents the beginning of a day, and the color red represents the recording of a "null" day.

Once the hdf5 file has been created, the desirable functions to use are:

    - norm_daily_data
    - norm_intraday_data
    
If the hdf5 file has not been created, then run the 'setup_hdf' function.
'''


import itertools
import numpy
import os
import pandas
import types
import xlrd


# Excel Color constants
BG_GREEN = 31
BG_RED = 10

# Filepath constants
DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data')
FILE_XL_INTAKE = os.path.join(DATA_DIR, 'retro_omeq_intake.xls')
FILE_XL_SCORES = os.path.join(DATA_DIR, 'retro_pain_scores.xls')
FILE_XL_DETAIL = os.path.join(DATA_DIR, 'patient features.csv')
FILE_RETRO = os.path.join(DATA_DIR, 'retro.h5')

# Sheet / key constants
SHEET_INTAKE = 'RETROSPECTIVE OMeq Intake'
SHEET_SCORES = 'RETROSPECTIVE Pain Scores'
KEY_INTRADAY = 'intraday'
KEY_DAILY = 'daily'
KEY_DETAIL = 'detail'


def cell_bg_color(cell: xlrd.sheet.Cell, wb: xlrd.Book) -> int:
    '''
    Extract the background color index of a given cell

    Args:
        cell:
            Cell, the cell to extract background information from
    
        wb:
            Book, the workbook

    Returns:
        int, the color index value
        >>> wb = xlrd.open_workbook(FILE_XL_INTAKE, formatting_info=True)
        >>> sheet = wb.sheet_by_name(SHEET_INTAKE)
        >>> cell = sheet.cell(0,0)
        >>> cell_bg_color(cell, wb)
        22
    '''
    return wb.xf_list[cell.xf_index].background.pattern_colour_index


def sheet_to_dataframe(path: str, sheet_name: str, cell_extractor: types.FunctionType, columns: list, num_rows: int=0, num_cols: int=0, start: tuple=(0,0)):
    '''
    Convert a sheet to a pandas.DataFrame by processing each cell with a cell extractor. This was necessary over pandas.read_excel due to the custom
    formatting of data to be extracted (color-coding, pivoted-shape, etc.).

    Args:
        path:
            str, the path to the file

        sheet_name:
            str, the name of the sheet in the file

        cell_extractor:
            Function, the function that accepts 4 arguments (row number, col number, cell object, workbook object) and returns a tuple of data

        columns:
            list[str] a list of column names

        num_rows:
            int, default 0, the number of rows to read

        num_cols:
            int, default 0, the number of columns to read

        start:
            tuple[int, int], default (0,0), the starting cell location

    Returns:
        DataFrame, the resulting data
        >>> sheet_to_dataframe(FILE_XL_INTAKE, SHEET_INTAKE, intake_extractor, ['Patient', 'DayNum', 'Intake'], 155, 28, (1,5)).head()
           Patient  DayNum  Intake
        0        1       1    80.0
        1        1       2    60.0
        2        1       3    70.0
        3        1       4    40.0
        4        2       1    75.0
    '''
    wb = xlrd.open_workbook(path, formatting_info=True)
    sheet = wb.sheet_by_name(sheet_name)
    data = []
    for row, col in itertools.product(range(start[0], start[0] + num_rows, 1), range(start[1], start[1] + num_cols, 1)):
            cell_data = cell_extractor(row, col, sheet.cell(row, col), wb)
            if cell_data is not None:
                data.append(cell_data)
    return pandas.DataFrame(data, columns=columns)


def intake_extractor(row: int, col: int, cell: xlrd.sheet.Cell, wb: xlrd.Book):
    '''
    Extract cell information from the intake sheet

    Args:
        row:
            int, row number

        col:
            int, columns number

        cell:
            Cell, the cell from which to extract data

        wb:
            Book, the workbook

    Returns:
        tuple (row, adjusted column number, cell value) if cell is not empty
    '''
    if not cell.value == '':
        return (row, col-4, cell.value)


def pain_score_extractor(row: int, col: int, cell: xlrd.sheet.Cell, wb: xlrd.Book):
    '''
        Extract cell information from the intake sheet

        Args:
            row:
                int, row number

            col:
                int, columns number

            cell:
                Cell, the cell from which to extract data

            wb:
                Book, the workbook

        Returns:
            tuple (row, adjusted column number, cell value) if cell is not empty
        '''
    if not cell.value == '':
        return (row, col-4, cell.value, int(cell_bg_color(cell, wb) == BG_GREEN))
    elif cell_bg_color(cell, wb) == BG_RED: # Red signifies an empty day
        return (row, col-4, numpy.nan, 1)


def load_intake_data():
    '''
    Load data from the intake spreadsheet

    Notes:
        Convenience:
            This function is a convenience wrapper around sheet_to_dataframe

    Returns:
        DataFrame, the retrospective intake data
        >>> load_intake_data().head()
           Patient  DayNum  Intake
        0        1       1    80.0
        1        1       2    60.0
        2        1       3    70.0
        3        1       4    40.0
        4        2       1    75.0
    '''
    return sheet_to_dataframe(path=FILE_XL_INTAKE,
                              sheet_name=SHEET_INTAKE,
                              cell_extractor=intake_extractor,
                              columns=['Patient', 'DayNum', 'Intake'],
                              num_rows=154, num_cols=28, start=(1, 5))


def load_scores_data():
    '''
    Load data from the pain scores spreadsheet

    Notes:
        Convenience:
            This function is a convenience wrapper around sheet_to_dataframe

    Returns:
        DataFrame, the retrospective scores data
        >>> load_scores_data().head()
           Patient  Ordinal  PainScore  DayStart  DayNum
        0        1        1        9.0         1       1
        1        1        2        2.0         0       1
        2        1        3        2.0         0       1
        3        1        4        8.0         0       1
        4        1        5        0.0         0       1
    '''
    df = sheet_to_dataframe(path=FILE_XL_SCORES,
                            sheet_name=SHEET_SCORES,
                            cell_extractor=pain_score_extractor,
                            columns=['Patient', 'Ordinal', 'PainScore', 'DayStart'],
                            num_rows=155, num_cols=202, start=(1, 5))
    df['DayNum'] = df[['Patient', 'DayStart']].groupby('Patient').cumsum()
    return df


def load_detail_data():
    '''
    Load data from the patient detail sheet

    Returns:
        DataFrame, the retrospective detail data
        >>> load_detail_data().head(1).T
                                               0
        Patient                                1
        AgeAtAdmit                            32
        Gender                            Female
        ImpairmentGroup  Spinal_Cord_Dysfunction
        Depression                             0
    '''
    df = pandas.read_csv(FILE_XL_DETAIL)
    return df.rename(columns={'Subject': 'Patient',
                              'Age at Admit': 'AgeAtAdmit',
                              'Impairment Group': 'ImpairmentGroup',
                              'Depression (1=Y, 0=N)': 'Depression'})


def intraday_data():
    '''
    Load a complete intraday dataset, inclusive of patient details

    Returns:
        DataFrame, the intraday data
        >>> intraday_data().head(1).T
                                               0
        Patient                                1
        Ordinal                                1
        PainScore                              9
        DayStart                               1
        DayNum                                 1
        Intake                                80
        AgeAtAdmit                            32
        Gender                            Female
        ImpairmentGroup  Spinal_Cord_Dysfunction
        Depression                             0
    '''
    scores = load_scores_data()
    intake = load_intake_data()
    details = load_detail_data()
    return scores.merge(intake, 'left', on=['Patient', 'DayNum']).merge(details, 'left', on='Patient')


def daily_data():
    '''
    Load a complete intraday dataset, inclusive of patient details

    Returns:
        DataFrame, the daily aggregated data
        >>> daily_data().head(1).T
                                               0
        Patient                                1
        DayNum                                 1
        Intake                                80
        PainScore                             29
        Ordinal                                7
        AgeAtAdmit                            32
        Gender                            Female
        ImpairmentGroup  Spinal_Cord_Dysfunction
        Depression                             0
    '''
    agg_funcs = {
        'Intake': lambda x: x.values[0],
        'PainScore': lambda x: x.sum(),
        'Ordinal': lambda x: x.count(),
        'AgeAtAdmit': lambda x: x.values[0],
        'Gender': lambda x: x.values[0],
        'ImpairmentGroup': lambda x: x.values[0],
        'Depression': lambda x: x.values[0],
    }
    data = intraday_data()
    return data.groupby(['Patient','DayNum']).agg(agg_funcs).reset_index().rename(columns={'Ordinal':'NumObs'})


def integrity_intraday_data():
    '''
    Check to make sure that the reported pain scores and intake data have the same number of days per patient

    Returns:
        DataFrame, patients with unequal painscore and intake days
        >>> integrity_intraday_data()
                 NumPainScoreDays  NumIntakeDays
        Patient                                 
        20                     10              9
        37                     16             17
        46                     22             23
        102                    12             13
        119                     9             10
        121                    13             14
        137                    11             10
        145                    23             22
    '''
    pain_scores = load_scores_data()
    intake = load_intake_data()
    day_counts = pandas.merge(
        pain_scores.groupby('Patient')[['DayNum']].max().rename(columns={'DayNum': 'NumPainScoreDays'}),
        intake.groupby('Patient')[['DayNum']].max().rename(columns={'DayNum': 'NumIntakeDays'}),
        left_index=True, right_index=True)
    return day_counts[day_counts.NumPainScoreDays != day_counts.NumIntakeDays]


def setup_hdf():
    '''
    Create an hdf5 file to store the normalized datasets
    >>> setup_hdf()
    '''
    intraday = intraday_data()
    daily = daily_data()
    detail = load_detail_data()
    intraday.to_hdf(FILE_RETRO, KEY_INTRADAY, format='table', mode='w')
    daily.to_hdf(FILE_RETRO, KEY_DAILY, format='table')
    detail.to_hdf(FILE_RETRO, KEY_DETAIL, format='table')


def norm_detail_data():
    '''
    Load normalized daily data from hdf5

    Returns:
        >>> norm_detail_data().head(1).T
                                               0
        Patient                                1
        DayNum                                 1
        Intake                                80
        PainScore                             29
        Ordinal                                7
        AgeAtAdmit                            32
        Gender                            Female
        ImpairmentGroup  Spinal_Cord_Dysfunction
        Depression                             0
    '''
    return pandas.read_hdf(FILE_RETRO, KEY_DETAIL)


def norm_daily_data():
    '''
    Load normalized daily data from hdf5

    Returns:
        >>> norm_daily_data().head(1).T
                                               0
        Patient                                1
        DayNum                                 1
        Intake                                80
        PainScore                             29
        Ordinal                                7
        AgeAtAdmit                            32
        Gender                            Female
        ImpairmentGroup  Spinal_Cord_Dysfunction
        Depression                             0
    '''
    return pandas.read_hdf(FILE_RETRO, KEY_DAILY)


def norm_intraday_data():
    '''
    Load normalized intraday data from hdf5

    Returns:
        >>> norm_intraday_data().head(1).T
                                               0
        Patient                                1
        Ordinal                                1
        PainScore                              9
        DayStart                               1
        DayNum                                 1
        Intake                                80
        AgeAtAdmit                            32
        Gender                            Female
        ImpairmentGroup  Spinal_Cord_Dysfunction
        Depression                             0
    '''
    return pandas.read_hdf(FILE_RETRO, KEY_INTRADAY)