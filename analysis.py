import pandas
import pandas as pd
from datetime import datetime


def latest_monthly():
    df = pd.read_csv('month.csv')

    month_dict = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6,
                  'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10, 'XI': 11, 'XII': 12}

    # Get latest data month
    year = (df.columns[-1].split(',')[0])
    month = month_dict[(df.columns[-1].split(',')[1])]  # month as integer from dict

    # Calculating same period change
    if df.columns[-1].split(',')[1] != 'I':
        from_year_start = df.iloc[:, -1:-month-1:-1].sum(axis=1)
        same_period = df.iloc[:, -1-12:-month-1-12:-1].sum(axis=1)
        change = (from_year_start/same_period-1) * 100
        abs_change = from_year_start - same_period
        df_label = df['English']
        df_change = pd.DataFrame({'label': df_label, 'change': change, 'absolute change': abs_change})
        df_change.to_csv('change.csv', index=False)

