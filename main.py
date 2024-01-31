import pandas as pd
import sys
import tkinter as tk
from tkinter import filedialog
import time

print("Select database files to analyse.")
time.sleep(5)

# create tkinter root window (it won't be shown)
root = tk.Tk()
root.withdraw()

# show file selection dialog for multiple files
selected_db_files = filedialog.askopenfilenames(title="Select DB files", filetypes=[("All Files", "*.xlsx")])

# check if user cancelled
if not selected_db_files:
    print("DB file selection cancelled.")
    sys.exit()
else:
    print(str(len(selected_db_files)) + " DB files selected.")

print("Select the Excel filters worksheet")
time.sleep(5)

# create tkinter root window (it won't be shown)
root = tk.Tk()
root.withdraw()

# show file selection dialog for multiple files
filepath_filter = filedialog.askopenfilename(title="Select the Excel filters worksheet",
                                             filetypes=[("All Files", "*.xlsx")])

# check if user cancelled
if not filepath_filter:
    print("Excel filters worksheet selection cancelled.")
    sys.exit()
else:
    print("Excel filters worksheet selected.")

# start looping through DB files
for file in selected_db_files:

    filepath_db = file

    print("Uploading " + filepath_db + " data. This will take a few minutes...")
    # import data
    df = pd.read_excel(filepath_db, engine='openpyxl')
    df_filter = pd.read_excel(filepath_filter)

    print("Upload completed.")

    table_1 = df_filter[df_filter.columns[0:3]].dropna(axis=0)
    table_2 = df_filter[df_filter.columns[4:7]].dropna(axis=0)
    table_3 = df_filter[df_filter.columns[8:12]].dropna(axis=0)

    # create column to hold time slots
    df['time_filter'] = df[df.columns[0]].astype(str).str[-5:]

    time_slots = df['time_filter'].unique()

    # dataframe to hold final results
    df_combined = pd.DataFrame()
    df_filtered = pd.DataFrame()

    df_filtered = pd.concat([df_filtered, df], ignore_index=True)


    def is_numeric(value):
        try:
            pd.to_numeric(value)
            return True
        except (ValueError, TypeError):
            return False

    # suppress warning
    pd.options.mode.chained_assignment = None

    # table_2 filter
    if len(table_2) > 0:
        for i in range(0, len(table_2)):
            # excel returns 32.0 etc
            # excel indexing starts from 1. adjustment
            filter_column_index = int(table_2.at[i, 'Table_2_column_index']) - 1
            filter_value_raw = table_2.at[i, 'Table_2_values']

            if is_numeric(filter_value_raw):
                print("Comparison operator missing from filter table. Process stopped 1.")
                sys.exit()
            elif filter_value_raw[-1] == "%":
                filter_value = float(filter_value_raw[1:-1]) / 100
            else:
                filter_value = float(filter_value_raw[1:])

            if ">" in filter_value_raw:
                df_filtered[df_filtered.columns[filter_column_index]] = pd.to_numeric(
                    df_filtered[df_filtered.columns[filter_column_index]], errors="coerce")
                df_filtered = df_filtered[(df_filtered[df_filtered.columns[filter_column_index]] > filter_value) & (
                    df_filtered[df_filtered.columns[filter_column_index]].notna())]

            elif "<" in filter_value_raw:
                df_filtered[df_filtered.columns[filter_column_index]] = pd.to_numeric(
                    df_filtered[df_filtered.columns[filter_column_index]], errors="coerce")
                df_filtered = df_filtered[(df_filtered[df_filtered.columns[filter_column_index]] < filter_value) & (
                    df_filtered[df_filtered.columns[filter_column_index]].notna())]

            elif "=" in filter_value_raw:
                df_filtered[df_filtered.columns[filter_column_index]] = pd.to_numeric(
                    df_filtered[df_filtered.columns[filter_column_index]], errors="coerce")
                df_filtered = df_filtered[(df_filtered[df_filtered.columns[filter_column_index]] == filter_value) & (
                    df_filtered[df_filtered.columns[filter_column_index]].notna())]
            else:
                pass
    else:
        pass

    # table 2 filter and selection of Top 5
    for element in time_slots:
        df_filtered_2 = df_filtered[df_filtered['time_filter'] == element]

        # table_1 filter
        if len(table_1) > 0:
            for i in range(0, len(table_1)):
                # excel returns 32.0 etc
                # excel indexing starts from 1. adjustment
                filter_column_index = int(table_1.at[i, 'Table_1_column_index']) - 1
                filter_value_raw = table_1.at[i, 'Table_1_values']

                if is_numeric(filter_value_raw):
                    print("Comparison operator missing from filter table. Process stopped 1.")
                    sys.exit()
                elif filter_value_raw[-1] == "%":
                    filter_value = float(filter_value_raw[1:-1]) / 100
                else:
                    filter_value = float(filter_value_raw[1:])

                if ">" in filter_value_raw:
                    df_filtered_2[df_filtered_2.columns[filter_column_index]] = pd.to_numeric(
                        df_filtered_2[df_filtered_2.columns[filter_column_index]], errors="coerce")
                    df_tf = df_filtered_2[(df_filtered_2[df_filtered_2.columns[filter_column_index]] > filter_value) & (
                        df_filtered_2[df_filtered_2.columns[filter_column_index]].notna())]

                    # Select the top 5 returns of "Column AZ, N+2 C vs. Close"
                    df_tf_sorted = df_tf.sort_values(by='N+2 C vs. Close', ascending=False)
                    df_combined = pd.concat([df_tf_sorted.head(), df_combined], ignore_index=True)

                elif "<" in filter_value_raw:
                    df_filtered_2[df_filtered_2.columns[filter_column_index]] = pd.to_numeric(
                        df_filtered_2[df_filtered_2.columns[filter_column_index]], errors="coerce")
                    df_tf = df_filtered_2[(df_filtered_2[df_filtered_2.columns[filter_column_index]] < filter_value) & (
                        df_filtered_2[df_filtered_2.columns[filter_column_index]].notna())]

                    # Select the top 5 returns of "Column AZ, N+2 C vs. Close"
                    df_tf_sorted = df_tf.sort_values(by='N+2 C vs. Close', ascending=False)
                    df_combined = pd.concat([df_tf_sorted.head(), df_combined], ignore_index=True)

                elif "=" in filter_value_raw:
                    df_filtered_2[df_filtered_2.columns[filter_column_index]] = pd.to_numeric(
                        df_filtered_2[df_filtered_2.columns[filter_column_index]], errors="coerce")
                    df_tf = df_filtered_2[
                        (df_filtered_2[df_filtered_2.columns[filter_column_index]] == filter_value) & (
                            df_filtered_2[df_filtered_2.columns[filter_column_index]].notna())]

                    # Select the top 5 returns of "Column AZ, N+2 C vs. Close"
                    df_tf_sorted = df_tf.sort_values(by='N+2 C vs. Close', ascending=False)
                    df_combined = pd.concat([df_tf_sorted.head(), df_combined], ignore_index=True)

                else:
                    print("Comparison operator missing from filter table. Process stopped 2.")
                    # sys.exit()
        else:
            df_combined = pd.concat([df_filtered_2, df_combined], ignore_index=True)

    # table 3 filter
    if len(table_3) > 0:
        for i in range(0, len(table_3)):
            # excel returns 32.0 etc
            # excel indexing starts from 1. adjustment
            filter_column_index = int(table_3.iloc[i, 1]) - 1
            filter_comparison = table_3.iloc[i, 2]
            filter_value_raw = table_3.iloc[i, 3]

            if is_numeric(filter_value_raw):
                filter_value = filter_value_raw
            elif filter_value_raw[-1] == "%":
                filter_value = float(filter_value_raw[:-1])
            else:
                print("Check filter values in table 3. Are they all numeric?")
                sys.exit()

            if filter_comparison == ">=":
                df_combined[df_combined.columns[filter_column_index]] = pd.to_numeric(
                    df_combined[df_combined.columns[filter_column_index]], errors="coerce")
                df_combined = df_combined[(df_combined[df_combined.columns[filter_column_index]] >= filter_value) & (
                    df_combined[df_combined.columns[filter_column_index]].notna())]

            elif filter_comparison == "<=":
                df_combined[df_combined.columns[filter_column_index]] = pd.to_numeric(
                    df_combined[df_combined.columns[filter_column_index]], errors="coerce")
                df_combined = df_combined[(df_combined[df_combined.columns[filter_column_index]] <= filter_value) & (
                    df_combined[df_combined.columns[filter_column_index]].notna())]

            else:
                print("Table 3 allows for >= and <= comparisons only.")
                sys.exit()

    # drop 'time_filter' column in df_combined
    df_combined.drop('time_filter', axis=1, inplace=True)

    # summary statistics
    df_summary = pd.DataFrame()
    df_summary.loc[0, 0] = "DB Statistics"
    df_summary.loc[0, 1] = "Returns of  N+2 C vs. Close on DB Average"
    df_summary.loc[1, 1] = df_combined['N+2 C vs. Close'].mean()
    df_summary.loc[0, 2] = "Returns of  N+2 C vs. Close on DB  Trade"
    df_summary.loc[1, 2] = df_combined['N+2 C vs. Close'].count()
    df_summary.loc[0, 3] = "Returns of  N+2 C vs. Close on DB  Total"
    df_summary.loc[1, 3] = df_combined['N+2 C vs. Close'].sum()
    df_summary.loc[0, 4] = "Returns of N+5 C vs. Close on DB Average"
    df_summary.loc[1, 4] = df_combined['N+5 C vs. Close'].mean()
    df_summary.loc[0, 5] = "Returns of N+5 C vs. Close on DB  Trade"
    df_summary.loc[1, 5] = df_combined['N+5 C vs. Close'].count()
    df_summary.loc[0, 6] = "Returns of N+5 C vs. Close on DB  Total"
    df_summary.loc[1, 6] = df_combined['N+5 C vs. Close'].sum()

    # path to save
    excel_path = filepath_db[:-5] + "_analysed.xlsx"

    print("Saving " + excel_path)

    # create Excel write
    with pd.ExcelWriter(excel_path, engine='xlsxwriter', mode='w') as writer:
        # save dataframe to an Excel file
        df_combined.to_excel(writer, sheet_name="db_data", index=False)
        df_summary.to_excel(writer, sheet_name="summary", index=False, header=False)

        # get the xlsxwriter workbook object
        workbook = writer.book

        # add a cell format with wrap text
        wrap_format = workbook.add_format({'text_wrap': True})

        # iterate through all worksheets
        for sheet_name in writer.sheets.keys():
            worksheet = writer.sheets[sheet_name]

            # set wrap text for all cells in the first row of each sheet
            worksheet.set_row(0, None, wrap_format)

        # Add number formats
        percent_format = workbook.add_format({'num_format': '0.00%'})
        percent_format_0 = workbook.add_format({'num_format': '0%'})
        percent_format_1 = workbook.add_format({'num_format': '0.0%'})
        number_format = workbook.add_format({'num_format': '0'})

        # set the formats
        worksheet = writer.sheets["summary"]

        worksheet.set_column('B:B', None, percent_format)
        worksheet.set_column('D:D', None, percent_format)
        worksheet.set_column('E:E', None, percent_format)
        worksheet.set_column('G:G', None, percent_format)

        worksheet = writer.sheets["db_data"]
        worksheet.set_column('D:D', None, percent_format_0)
        worksheet.set_column('E:E', None, percent_format_0)
        worksheet.set_column('I:I', None, percent_format_0)
        worksheet.set_column('J:J', None, percent_format_0)
        worksheet.set_column('K:K', None, percent_format_0)
        worksheet.set_column('L:L', None, percent_format_0)
        worksheet.set_column('S:S', None, percent_format_0)
        worksheet.set_column('T:T', None, percent_format_0)
        worksheet.set_column('U:U', None, percent_format_0)
        worksheet.set_column('X:X', None, percent_format_0)
        worksheet.set_column('AB:AB', None, percent_format_0)
        worksheet.set_column('AF:AF', None, percent_format_1)
        worksheet.set_column('AG:AG', None, percent_format_1)
        worksheet.set_column('AH:AH', None, percent_format_1)
        worksheet.set_column('AI:AI', None, percent_format_0)
        worksheet.set_column('AJ:AJ', None, percent_format_0)
        worksheet.set_column('AK:AK', None, percent_format_0)
        worksheet.set_column('AL:AL', None, percent_format_0)
        worksheet.set_column('AO:AO', None, percent_format_0)
        worksheet.set_column('AR:AR', None, percent_format_1)
        worksheet.set_column('BB:BB', None, percent_format_1)
        worksheet.set_column('BC:BC', None, percent_format_1)
        worksheet.set_column('BD:BD', None, percent_format_1)
        worksheet.set_column('BE:BE', None, percent_format_1)
        worksheet.set_column('BF:BF', None, percent_format_1)

    print("Saved successfully.")

# reset warning option to its default
pd.options.mode.chained_assignment = 'warn'

print("Job completed successfully.")
