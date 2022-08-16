#! /bin/python3

import pandas as pd
import datetime as dt
import calendar

KAKEIBO_DIR = r"C:\Users\nishi\Dropbox\家計簿\household_account\Dashboard.xlsx"
OUT_DIR = r"C:\Users\nishi\Dropbox\家計簿\household_account\TEST.xlsx"
BEGIN_YEAR = 2019
END_YEAR = 2022


def update_monthly_IoE(year, month):
    begging_month = 1
    end_of_month = calendar.monthrange(year, month)[1]
    df_month = df_in[(df_in['日付'] > dt.datetime(year, month, begging_month)) &
                     (df_in['日付'] < dt.datetime(year, month, end_of_month))]
    
    categolies = get_categolies(df_month)

    sr_month_IoE = pd.Series(
        index=categolies,
        dtype='float64',
        name='%s-%s' % (year, month)
    )

    for categoly in categolies:
        pri_item, sec_item = categoly.split('-')
        total_IoE = df_month[df_month['大項目'].str.cat(
            df_month['中項目'], sep='-') == categoly]['金額（円）'].sum()
        sr_month_IoE[categoly] = pd.Series(total_IoE)
    
    return sr_month_IoE


def get_categolies(df):
    categolies = list(set(df['大項目'].str.cat(df['中項目'], sep='-').tolist()))
    return categolies


if __name__ == '__main__':
    df_in = pd.read_excel(KAKEIBO_DIR, sheet_name='data')
    df_in['日付'] = pd.to_datetime(df_in['日付'])
    df_out = pd.read_excel(KAKEIBO_DIR, sheet_name='dashboard')

    sr_month_IoE = pd.DataFrame(dtype='float64')
    for year in range(BEGIN_YEAR, END_YEAR):
        for month in range(1, 12):
            sr_monthly = update_monthly_IoE(year, month)
            sr_month_IoE = pd.concat(
                [sr_month_IoE, sr_monthly], join='outer', axis=1
            )

    with pd.ExcelWriter(OUT_DIR) as w:
        sr_month_IoE.to_excel(
            w, sheet_name='dashboard',
            index=True, header=True
        )
