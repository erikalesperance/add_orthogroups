import os
import pandas as pd

def process_files(directory, output_filename):
    """Reads all .csv files in a directory, adds Arabidopsis orthogroups, and outputs them as a single Excel file 
    with different sheets."""
    
    # Read orthogroups data and normalize column names
    orthogroups = pd.read_csv("Orthogroups_Asterales_bidensv5soft.csv", sep=",")  # Change sep to ","
    orthogroups.columns = orthogroups.columns.str.strip().str.lower()  # Normalize column names
    
    print(orthogroups.columns)  # Print the columns to verify

    column_to_search = 'bidensv5_hap1_soft_transdecoder'.lower()  # Match normalized column name
    if column_to_search not in orthogroups.columns:
        print(f"Error: Column '{column_to_search}' not found in the orthogroups file.")
        return
    
    dataframes = {}
    values_lists = []

    for filename in os.listdir(directory):
        if filename.endswith('.csv'):
            filepath = os.path.join(directory, filename)
            df = pd.read_csv(filepath)  # Read the input CSV file
            df_subset = df[df['padj'] < 0.05]  # Filter DataFrame based on a condition
            dataframes[filename] = df_subset

            values_list = []
            for index, row in df_subset.iterrows():
                # Attempt to access 'Unnamed: 0' or use the first column if it doesn't exist
                try:
                    value_to_check = row['Unnamed: 0']
                except KeyError:
                    value_to_check = row.iloc[0]  # Use the first column if 'Unnamed: 0' is not found

                value_found = False
                for separate_index, separate_row in orthogroups.iterrows():
                    try:
                        if pd.notna(separate_row[column_to_search]) and value_to_check in separate_row[column_to_search]:
                            separate_value = separate_row['arabidopsis']  # Use normalized name
                            values_list.append(separate_value)
                            value_found = True
                            break
                    except TypeError:
                        values_list.append('nan')
                if not value_found:
                    values_list.append('nan')
            values_lists.append(values_list)

    new_column_name = 'Arabidopsis_orthologs'

    for df, values_list in zip(dataframes.values(), values_lists):
        df[new_column_name] = values_list

    excel_writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')

    for df_name, df in dataframes.items():
        df.to_excel(excel_writer, sheet_name=df_name, index=False)

   
    excel_writer.close()

# Process files for I v II
process_files("IvII/", "I_v_II_AtID.xlsx")

# Process files for I v III
process_files("IvIII/", "I_v_III_AtID.xlsx")

# Process files for I v IV
process_files("IvIV/", "I_v_IV_AtID.xlsx")

# Process files for II v III
process_files("IIvIII/", "II_v_III_AtID.xlsx")

# Process files for II v IV
process_files("IIvIV/", "II_v_IV_AtID.xlsx")

# Process files for III v IV
process_files("IIIvIV/", "III_v_IV_AtID.xlsx")
