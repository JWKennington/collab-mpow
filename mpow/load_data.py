'''This module contains utilities for loading and parsing data for MPOW. The functions below parse data from Excel sheets (xls format). The 
data are arranged such that each row represents the time-ordered observations for a particular patient. In the pain score data, the 
color green represents the beginning of a day, and the color red represents the recording of a "null" day.

Once the hdf5 file has been created, the desirable functions to use are:

    - norm_daily_data
    - norm_intraday_data
    
If the hdf5 file has not been created, then run the 'setup_hdf' function.
'''


import collections
import itertools
import numpy
import os
import pandas
import pathlib
import types
import xlrd


# Excel Color constants
BG_GREEN = 31
BG_ORANGE = 29
BG_RED = 10


# Filepath constants
DATA_DIR = pathlib.Path(__file__).parent.parent / 'data'
FILE_XL_INTAKE = DATA_DIR / 'retro_omeq_intake.xls'
FILE_XL_SCORES = DATA_DIR / 'retro_pain_scores.xls'
FILE_XL_DETAIL = DATA_DIR / 'patient features.csv'
FILE_RETRO = DATA_DIR / 'retro.h5'
FILE_PROTO = DATA_DIR / 'MPOW-Data-20190206_xls.xls'
FILE_HDF = DATA_DIR / 'data.h5'


# Sheet / key constants
SHEET_INTAKE = 'OMeq Intake'
SHEET_SCORES = 'Pain Scores'
KEY_INTRADAY = 'intraday'
KEY_DAILY = 'daily'
KEY_DETAIL = 'detail'


Source = collections.namedtuple('Source', 'file page_prefix')
class Sources:
    Retrospective = Source(FILE_PROTO.as_posix(), 'RETROSPECTIVE')
    Protocol = Source(FILE_PROTO.as_posix(), 'PROTOCOL')


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
        >>> sheet = wb.sheet_by_name(RETRO_SHEET_INTAKE)
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
        >>> sheet_to_dataframe(FILE_XL_INTAKE.as_posix(), Sources.Retrospective.page_prefix + ' ' + SHEET_INTAKE, intake_extractor, ['Patient', 'DayNum', 'Intake'], 150, 28, (1,5)).head()
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
    '''Extract cell information from the intake sheet

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
    '''Extract cell information from the intake sheet

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
    cell_color = cell_bg_color(cell, wb) 
    if not cell.value == '':
        return (row, col-15, cell.value, int(cell_color == BG_ORANGE))
    elif cell_color == BG_RED: # Red signifies an empty day
        return (row, col-15, numpy.nan, 1)


def load_intake_data(source: Source=Sources.Retrospective):
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

        >>> load_intake_data(source=Sources.Protocol).head()
           Patient  DayNum  Intake
        0        1       1     5.0
        1        1       2     0.0
        2        1       3     0.0
        3        1       4     0.0
        4        1       5     0.0
    '''
    return sheet_to_dataframe(path=source.file,
                              sheet_name=source.page_prefix + ' ' + SHEET_INTAKE,
                              cell_extractor=intake_extractor,
                              columns=['Patient', 'DayNum', 'Intake'],
                              num_rows=154, num_cols=28, start=(1, 5))


def load_scores_data(source: Source=Sources.Retrospective):
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

        >>> load_scores_data(source=Sources.Protocol).head()
           Patient  Ordinal  PainScore  DayStart  DayNum
        0        1        1        0.0         1       1
        1        1        2        0.0         0       1
        2        1        3        0.0         0       1
        3        1        4        7.0         0       1
        4        1        5        0.0         0       1
    '''
    df = sheet_to_dataframe(path=source.file,
                            sheet_name=source.page_prefix + ' ' + SHEET_SCORES,
                            cell_extractor=pain_score_extractor,
                            columns=['Patient', 'Ordinal', 'PainScore', 'DayStart'],
                            num_rows=155, num_cols=202, start=(1, 16))
    df['DayNum'] = df[['Patient', 'DayStart']].groupby('Patient').cumsum()
    return df


def load_detail_data(source: Source=Sources.Retrospective):
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

        >>> load_detail_data(source=Sources.Protocol).head(1).T
                                0
        Patient                 1
        AgeAtAdmit             78
        Gender               Male
        ImpairmentGroup  Debility
        Depression              1
    '''
    if source == Sources.Retrospective:
        df = pandas.read_csv(FILE_XL_DETAIL)
    elif source == Sources.Protocol:
        df = pandas.read_excel(source.file, sheetname=source.page_prefix + ' ' + 'Data collection')
    else:
        raise ValueError('Unknown source: {}'.format(source))
    df = df.rename(columns={'Subject': 'Patient',
                            'Age at Admit': 'AgeAtAdmit',
                            'Impairment Group': 'ImpairmentGroup',
                            'Depression (1=Y, 0=N)': 'Depression'})
    return df[['Patient', 'AgeAtAdmit', 'Gender', 'ImpairmentGroup', 'Depression'].copy()]


def intraday_data(source: Source=Sources.Retrospective):
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

        >>> intraday_data(source=Sources.Protocol).head(1).T
                                0
        Patient                 1
        Ordinal                 1
        PainScore               0
        DayStart                1
        DayNum                  1
        Intake                  5
        AgeAtAdmit             78
        Gender               Male
        ImpairmentGroup  Debility
        Depression              1
    '''
    scores = load_scores_data(source=source)
    intake = load_intake_data(source=source)
    details = load_detail_data(source=source)
    return scores.merge(intake, 'left', on=['Patient', 'DayNum']).merge(details, 'left', on='Patient')


def daily_data(source: Source=Sources.Retrospective):
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
        NumObs                                 7
        AgeAtAdmit                            32
        Gender                            Female
        ImpairmentGroup  Spinal_Cord_Dysfunction
        Depression                             0

        >>> daily_data(source=Sources.Protocol).head(1).T
                                0
        Patient                 1
        DayNum                  1
        Intake                  5
        PainScore               7
        NumObs                  6
        AgeAtAdmit             78
        Gender               Male
        ImpairmentGroup  Debility
        Depression              1
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
    data = intraday_data(source=source)
    return data.groupby(['Patient','DayNum']).agg(agg_funcs).reset_index().rename(columns={'Ordinal':'NumObs'})


def integrity_intraday_data(source: Source=Sources.Retrospective):
    '''
    Check to make sure that the reported pain scores and intake data have the same number of days per patient

    Returns:
        DataFrame, patients with unequal painscore and intake days
        >>> integrity_intraday_data()
                 NumPainScoreDays  NumIntakeDays
        Patient                                 
        20                      8              9

        >>> integrity_intraday_data(source=Sources.Protocol)
                 NumPainScoreDays  NumIntakeDays
        Patient                                 
        4                       9              8
        11                      7              8
        15                     11             12
        21                     15             16
        25                      6              7
        34                     18             16
        40                     15             16
        49                      6              7
        50                      1              9
        51                      1             13
    '''
    pain_scores = load_scores_data(source=source)
    intake = load_intake_data(source=source)
    day_counts = pandas.merge(
        pain_scores.groupby('Patient')[['DayNum']].max().rename(columns={'DayNum': 'NumPainScoreDays'}),
        intake.groupby('Patient')[['DayNum']].max().rename(columns={'DayNum': 'NumIntakeDays'}),
        left_index=True, right_index=True)
    return day_counts[day_counts.NumPainScoreDays != day_counts.NumIntakeDays]


def _setup_hdf(source: Source=Sources.Retrospective):
    '''Create an hdf5 file to store the normalized datasets
    '''
    
    intraday = intraday_data(source=source)
    daily = daily_data(source=source)
    detail = load_detail_data(source=source)
    intraday.to_hdf(FILE_HDF.as_posix(), source.page_prefix.lower() + '_' + KEY_INTRADAY, format='table')
    daily.to_hdf(FILE_HDF.as_posix(), source.page_prefix.lower() + '_' + KEY_DAILY, format='table')
    detail.to_hdf(FILE_HDF.as_posix(), source.page_prefix.lower() + '_' + KEY_DETAIL, format='table')


def setup_hdf():
    """Setup HDF file

    >>> setup_hdf()
    """
    if FILE_HDF.exists():
        os.remove(FILE_HDF.as_posix())
    _setup_hdf(source=Sources.Retrospective)
    _setup_hdf(source=Sources.Protocol)


def norm_detail_data(source: Source=Sources.Retrospective):
    '''
    Load normalized daily data from hdf5

    Returns:
        >>> norm_detail_data().head(1).T
                                               0
        Patient                                1
        AgeAtAdmit                            32
        Gender                            Female
        ImpairmentGroup  Spinal_Cord_Dysfunction
        Depression                             0

        >>> norm_detail_data(source=Sources.Protocol).head(1).T
                                0
        Patient                 1
        AgeAtAdmit             78
        Gender               Male
        ImpairmentGroup  Debility
        Depression              1
    '''
    return pandas.read_hdf(FILE_HDF, source.page_prefix.lower() + '_' + KEY_DETAIL)


def norm_daily_data(source: Source=Sources.Retrospective):
    '''
    Load normalized daily data from hdf5

    Returns:
        >>> norm_daily_data().head(1).T
                                               0
        Patient                                1
        DayNum                                 1
        Intake                                80
        PainScore                             29
        NumObs                                 7
        AgeAtAdmit                            32
        Gender                            Female
        ImpairmentGroup  Spinal_Cord_Dysfunction
        Depression                             0

        >>> norm_daily_data(source=Sources.Protocol).head(1).T
                                0
        Patient                 1
        DayNum                  1
        Intake                  5
        PainScore               7
        NumObs                  6
        AgeAtAdmit             78
        Gender               Male
        ImpairmentGroup  Debility
        Depression              1
    '''
    return pandas.read_hdf(FILE_HDF, source.page_prefix.lower() + '_' + KEY_DAILY)


def norm_intraday_data(source: Source=Sources.Retrospective):
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

        >>> norm_intraday_data(source=Sources.Protocol).head(1).T
                                0
        Patient                 1
        Ordinal                 1
        PainScore               0
        DayStart                1
        DayNum                  1
        Intake                  5
        AgeAtAdmit             78
        Gender               Male
        ImpairmentGroup  Debility
        Depression              1
    '''
    return pandas.read_hdf(FILE_HDF, source.page_prefix.lower() + '_' + KEY_INTRADAY)
