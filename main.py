import pandas as pd
import sys

filepath_db = "C:/Users/Bonbon/Documents/UP/Raymond Chan/New folder/DB_To_Jimmy/DB_To_Jimmy/v8.26_DB1_240128_v6_FF.xlsx"

filepath_filter = "C:/Users/Bonbon/Documents/UP/Raymond Chan/New folder/filter_table.xlsx"

# import data
df = pd.read_excel(filepath_db)
df_filter = pd.read_excel(filepath_filter)

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
                df_tf = df_tf.sort_values(by='N+2 C vs. Close', ascending=False)
                df_combined = pd.concat([df_tf.head(), df_combined], ignore_index=True)

            elif "<" in filter_value_raw:
                df_filtered_2[df_filtered_2.columns[filter_column_index]] = pd.to_numeric(
                    df_filtered_2[df_filtered_2.columns[filter_column_index]], errors="coerce")
                df_tf = df_filtered_2[(df_filtered_2[df_filtered_2.columns[filter_column_index]] < filter_value) & (
                    df_filtered_2[df_filtered_2.columns[filter_column_index]].notna())]

                # Select the top 5 returns of "Column AZ, N+2 C vs. Close"
                df_tf = df_tf.sort_values(by='N+2 C vs. Close', ascending=False)
                df_combined = pd.concat([df_tf.head(), df_combined], ignore_index=True)

            elif "=" in filter_value_raw:
                df_filtered_2[df_filtered_2.columns[filter_column_index]] = pd.to_numeric(
                    df_filtered_2[df_filtered_2.columns[filter_column_index]], errors="coerce")
                df_tf = df_filtered_2[(df_filtered_2[df_filtered_2.columns[filter_column_index]] == filter_value) & (
                    df_filtered_2[df_filtered_2.columns[filter_column_index]].notna())]

                # Select the top 5 returns of "Column AZ, N+2 C vs. Close"
                df_tf = df_tf.sort_values(by='N+2 C vs. Close', ascending=False)
                df_combined = pd.concat([df_tf.head(), df_combined], ignore_index=True)

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
        filter_column_index = int(table_3.at[i, 'Table_3_column_index']) - 1
        filter_comparison = table_3.at[i, 'Table_3_comparison']
        filter_value_raw = table_3.at[i, 'Table_3_values']

        if is_numeric(filter_value_raw):
            filter_value = filter_value_raw
        elif filter_value_raw[-1] == "%":
            filter_value = float(filter_value_raw[:-1]) / 100
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


