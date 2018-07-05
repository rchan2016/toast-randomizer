"""
Read calendar to extract the Thursdays dates for Toastmaster.
"""
import logging
import os
import pathlib
import sys

import numpy as np
import pandas as pd

MEMBER_LIST = pathlib.Path(os.getcwd(), 'member list.txt')
CALENDAR = pathlib.Path(os.getcwd(), '12-Month Calendar.xlsx')

MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
          'September', 'October', 'November', 'December']


def _read_calendar_excel_file():
    """

    :return:
    """
    month_calendar = pd.DataFrame()
    members = pd.read_csv(MEMBER_LIST, header=None)
    members.columns = ['Names']

    for month in MONTHS:
        df = pd.read_excel(CALENDAR, sheetname=month, parse_cols='B:I', header=1)
        df = df[[month.upper(), 'THU']]
        df['THU'] = df['THU'].dt.date
        df.dropna(subset=['THU'], inplace=True)
        header = [np.array([month] * len(df['THU'])),
                  np.array(df['THU'])]
        one_month = pd.DataFrame(index=members['Names'].tolist(), columns=header)
        month_calendar = pd.concat([month_calendar, one_month], axis=1)

    return month_calendar


def main():
    """
    Main module to read in the calendar.

    Returns:
         df (DataFrame): Containing months, Thursday dates and members name.
    """
    df = _read_calendar_excel_file()

    return df


if __name__ == '__main__':
    try:
        main()
    except Exception:
        logging.exception('{} failed'.format(__file__))
        sys.exit(1)
