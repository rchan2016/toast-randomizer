"""
Main module to create an excel report with meeting roles filled for each Thursday's
Toastmasters meeting.  Individual roles will be assigned randomly.
"""
import logging
import sys
import random

import numpy as np
import pandas as pd

import read_calendar

CLUB_ROSTER = 'C:/Users/rchan/PycharmProjects/Toast/club roster.xlsx'

ROLES = {'C': 'RED',
         'G': 'GREEN',
         'SP1': 'LIGHT BROWN',
         'SE1': 'YELLOW',
         'SP2': 'LIGHT RED',
         'SE2': 'ORANGE',
         'TT': 'BLUE',
         'GE': 'PINK',
         'T': 'LIGHT BLUE',
         'UM': 'PURPLE'}


def _add_roles(df):
    """

    :param df:
    :return:
    """
    df = df.copy()
    df = df.replace(np.nan, '', regex=True)
    count = dict()

    for person in df.index:
        count[person] = 0


    all_names = list(df.index).copy()
    remaining = all_names.copy()

    all_roles = list(ROLES).copy()

    for role in all_roles:

        previous_col = 0
        next_col = 1

        for col in range(len(df.columns)):
            if len(remaining) == 0:
                remaining = all_names.copy()
                remaining = random.sample(remaining, len(remaining))

            select_member = random.sample(remaining, 1)

            repeat = 0
            while True:
                if df.loc[select_member[0], df.columns[col]] == '':
                    if df.loc[select_member[0], df.columns[previous_col]] == '':
                        if df.loc[select_member[0], df.columns[next_col]] == '':
                            break
                if repeat == 27:
                    remaining = random.sample(
                        list(set(all_names) - set(select_member)),
                        len(all_names) - len(select_member))
                    repeat = 0
                repeat += 1
                select_member = random.sample(remaining, 1)

            df.loc[select_member[0], df.columns[col]] = role
            previous_col = col
            if next_col == 59:
                next_col = 0
            else:
                next_col = next_col + 1

            count[select_member[0]] +=1
            remaining.remove(select_member[0])

    print(count)

    return df

def main():
    """
    Main module to create the meeting role file.
    """
    df = read_calendar.main()

    df = _add_roles(df)

    sw_writer = pd.ExcelWriter(CLUB_ROSTER, engine='xlsxwriter',
                               date_format = 'yyyy-mm-dd', datetime_format='yyyy-mm-dd')
    df.to_excel(sw_writer, sheet_name='Club Roster')


if __name__ == '__main__':
    try:
        main()
    except Exception:
        logging.exception('{} failed'.format(__file__))
        sys.exit(1)
