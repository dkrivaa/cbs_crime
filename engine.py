from openpyxl import *
import pandas as pd
import streamlit as st
from io import BytesIO
import requests
import xlrd
from datetime import datetime
import re


def get_data():
    # GENERAL - Ranges to read data from
    data_ranges1 = [(11, 17), (19, 29), (30, 31), (33, 39), (41, 45), (47, 51), (52, 54), (64, 68),
                   (70, 82), (84, 87), (88, 89), (90,91), (92, 93)]

    data_ranges2 = [(11, 17), (19, 29), (30, 31), (33, 39), (41, 44), (46, 50), (51, 53), (63, 67),
                   (69, 81), (83, 86), (87, 88), (89,90), (91, 92)]

    data_ranges3 = [(11, 17), (19, 30), (31, 32), (34, 40), (42, 46), (48, 52), (53, 55), (65, 69),
                   (71, 83), (85, 88), (89, 90), (91, 92), (93, 94)]


    # Getting the text columns (english and hebrew) for all data
    url_text = 'https://www.cbs.gov.il/he/publications/doclib/2023/yarhon1123/q1.xls'

    response = requests.get(url_text)
    text_data = response.content

    # Open the .xls file using xlrd
    workbook = xlrd.open_workbook(file_contents=text_data)
    sheet = workbook.sheet_by_index(0)

    text_data = []
    for ranges in data_ranges3:
        text_data.extend([(sheet.cell_value(row, 5).strip(),
                           sheet.cell_value(row, 0).strip()) for row in range(ranges[0], ranges[1])])

    # Create a DataFrame from the list of values for text columns
    df = pd.DataFrame(text_data, columns=['Hebrew', 'English'])

    # Getting the actual data from each monthly bulletins
    month = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    # Initialize an empty list to store dataframes for each month
    dfs = []

    # Get the current year
    current_year = datetime.now().year % 100 + 1

    for year in range(16, current_year):
        for m in month:

            if year == 16 and m in ['01', '02', '03']:
                url = f'https://www.cbs.gov.il/he/publications/doclib/20{year}/yarhon{m}{year}/excel/q1.xls'
            elif year == 17 and m == month[6]:
                url = f'https://www.cbs.gov.il/he/publications/doclib/2018/yarhon0718/q1.xls'
            elif year < 21 and m == month[-1]:
                url = f'https://www.cbs.gov.il/he/publications/doclib/20{year + 1}/yarhon{m}{year}/q1.xls'
            else:
                url = f'https://www.cbs.gov.il/he/publications/doclib/20{year}/yarhon{m}{year}/q1.xls'

            try:
                response = requests.get(url)
                excel_data = response.content

                # Open the .xls file using xlrd
                workbook = xlrd.open_workbook(file_contents=excel_data)
                sheet = workbook.sheet_by_index(0)

                if year == 17 and m == month[6]:
                    title = f'{int(sheet.cell_value(3, 3))}' + ', ' + f'{sheet.cell_value(4, 3)}'
                else:
                    title = f'{int(sheet.cell_value(3, 1))}' + ', ' + f'{sheet.cell_value(4, 1)}'


                data1 = []
                data2 = []
                data3 = []
                if year < 18 or (year == 18 and m in ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11']) or (
                        year == 19 and m == month[-1]) or (year == 20 and m in ['01', '02', '03']):
                    for ranges in data_ranges1:
                        if year == 17 and m == month[6]:
                            data1.extend([sheet.cell_value(row, 4) for row in range(ranges[0], ranges[1])])
                        else:
                            data1.extend([sheet.cell_value(row, 2) for row in range(ranges[0], ranges[1])])

                    # Create DataFrame with a single column named based on the title
                    df_data1 = pd.DataFrame({title: data1})

                    # Adding empty row for 'Spread of sexually transmitted and other diseases' that was added 4/2020
                    # Create an empty row with 0
                    zero_row = pd.DataFrame([[0] * len(df_data1.columns)], columns=df_data1.columns)

                    # Use loc to insert the empty row at the specified index
                    df_data1 = pd.concat([df_data1.loc[:15 - 1], zero_row, df_data1.loc[15:]]).reset_index(
                        drop=True)

                    # Append df_data1 to the list
                    dfs.append(df_data1)

                elif (year == 18 and m == month[-1]) or (year == 19 and m in ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11']):
                    for ranges in data_ranges2:
                        data2.extend([sheet.cell_value(row, 2) for row in range(ranges[0], ranges[1])])
                    # Create DataFrame with a single column named based on the title
                    df_data2 = pd.DataFrame({title: data2})

                    # Adding 2 empty rows for 'Spread of sexually transmitted and other diseases' that was added 4/2020
                    # Create an empty row with 0
                    zero_row = pd.DataFrame([[0] * len(df_data2.columns)], columns=df_data2.columns)
                    # Use loc to insert the first empty row at the specified index
                    df_data2 = pd.concat([df_data2.loc[:15 - 1], zero_row, df_data2.loc[15:]]).reset_index(
                        drop=True)

                    # Create a new empty row with 0 for the second insertion
                    zero_row = pd.DataFrame([[0] * len(df_data2.columns)], columns=df_data2.columns)
                    # Use loc to insert the second empty row at the specified index
                    df_data2 = pd.concat([df_data2.loc[:27 - 1], zero_row, df_data2.loc[27:]]).reset_index(
                        drop=True)

                    # Append df_data to the list
                    dfs.append(df_data2)

                else:
                    for ranges in data_ranges3:
                        data3.extend([sheet.cell_value(row, 2) for row in range(ranges[0], ranges[1])])
                    # Create DataFrame with a single column named based on the title
                    df_data3 = pd.DataFrame({title: data3})

                    # Append df_data to the list
                    dfs.append(df_data3)

            except xlrd.biffh.XLRDError as xlrd_error:
                pass

    df_temp = pd.concat(dfs, axis=1, ignore_index=False)
    df = pd.concat([df, df_temp], axis=1, ignore_index=False)

    # Correcting the row titles
    df.at[17, 'Hebrew'] = 'כלפי חיי אדם'
    df.at[17, 'English'] = "Against a person's life"

    df.at[31, 'Hebrew'] = 'גידול, ייצור והפקת סמים'
    df.at[31, 'English'] = "Growth, manufacture and production of illicit drugs"

    df.at[37, 'Hebrew'] = 'התפרצות (לבית, לבית עסק או למוסד)'
    df.at[37, 'English'] = 'Breaking and entering (a home, business or institution)'

    df.at[53, 'Hebrew'] = 'סך הכל עבירות כלכליות'
    df.at[53, 'English'] = 'Total economic offences'

    df.at[54, 'Hebrew'] = 'סך הכל עבירות רישוי'
    df.at[54, 'English'] = 'Total licensing offences'

    df.at[55, 'Hebrew'] = 'סך הכל עבירות אחרות'
    df.at[55, 'English'] = 'Total other offences'

    # Saving files
    df.to_excel('RawData.xlsx', index=False)
    df.to_csv('RawData.csv', index=False)


def year_data():

    df = pd.read_csv('RawData.csv')

    # Making annual dataframe
    # Generate a list of column indices to drop where the last three characters are not 'Xll'
    columns_to_drop_year = [i for i in range(2, len(df.columns)) if df.columns[i][-3:] != 'XII']

    # Drop the selected columns and make df_year dataframe
    df_year = df.drop(df.columns[columns_to_drop_year], axis=1)
    df_year.to_excel('year.xlsx', index=False)
    df_year.to_csv('year.csv', index=False)
    # _________________________

def month_data():
    # Making monthly dataframe
    df = pd.read_csv('RawData.csv')

    my_list = []

    for i in range(len(df.columns) - 1, 2, -1):
        if df.columns[i].split(' ')[1] == 'I':
            sub_list = df[df.columns[i]].values
        else:
            sub_list = df[df.columns[i]].values - df[df.columns[i - 1]].values
        my_list.append(sub_list)

    # Reverse the list
    my_list.reverse()

    # Transpose the list of lists to create a DataFrame
    df_month = pd.DataFrame(my_list).transpose().fillna(0)

    # add row titles (hebrew and english)
    df_rows_titles = df.drop(df.columns[2:], axis=1)
    df_month = pd.concat([df_rows_titles, df_month], axis=1, ignore_index=False)

    # Making column titles
    # Split the string by both space and hyphen
    result = df.columns.str.split(r'\s+|-')
    result = result[3:]

    result = [x[0] + x[1] if len(x) == 2 else x[0] + x[2] for x in result]

    # Change column names starting from the specified index
    df_month.columns = list(df_month.columns[:2]) + result

    # Change the name of the specified column
    df_month.rename(columns={'2019,': '2019,I'}, inplace=True)

    df_month.to_excel('month.xlsx', index=False)
    df_month.to_csv('month.csv', index=False)



