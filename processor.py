import os
import re
import pandas as pd


def read_all_files_in_folder(folder_path):
    """
    Reads all supported files (CSV, Excel, JSON) from the given folder.

    Args:
        folder_path (str): Path to the folder containing files.

    Returns:
        list: A list of tuples (file_name, dataframe) for each file found.
    """
    dataframes = []

    if not os.path.exists(folder_path):
        print(f"Folder not found: {folder_path}")
        return dataframes

    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)

        if not os.path.isfile(file_path):
            continue

        ext = os.path.splitext(file_name)[1].lower()

        try:
            if ext == ".csv":
                df = pd.read_csv(file_path)
            elif ext in [".xls", ".xlsx"]:
                df = pd.read_excel(file_path, sheet_name=None)
            elif ext == ".json":
                df = pd.read_json(file_path)
            else:
                print(f"Unsupported file format: {file_name}")
                continue

            dataframes.append((file_name, df))
        except Exception as e:
            print(f"Error reading {file_name}: {e}")

    return dataframes


def extract_data_from_description(df):
    """
    Extracts structured information from the 'Description' column using regex patterns.

    Args:
        df (pd.DataFrame): Input dataframe containing a 'Description' column.

    Returns:
        pd.DataFrame: Dataframe with extracted columns added.
    """
    if "Description" not in df.columns:
        return df

    patterns = {
        "Category": r"Category:\s*(.*)",
        "Source": r"Source:\s*(.*)",
        "Client ID": r"Client ID:\s*(.*)",
        "Desk Ticket": r"Desk Ticket:\s*#(\d+)"
    }

    for col in patterns.keys():
        df[col] = ""

    for idx, row in df.iterrows():
        desc = str(row["Description"])
        for key, pattern in patterns.items():
            match = re.search(pattern, desc)
            if match:
                df.at[idx, key] = match.group(1).strip()

    return df


def process_all_files(input_folder, output_folder):
    """
    Processes all files in the input folder, extracts data, and saves results to output folder.

    Args:
        input_folder (str): Path to the folder containing input files.
        output_folder (str): Path to save processed output files.

    Returns:
        None
    """
    os.makedirs(output_folder, exist_ok=True)
    files = read_all_files_in_folder(input_folder)

    if not files:
        print("No files found to process")
        return

    for file_name, data in files:
        try:
            if isinstance(data, dict):  # Excel with multiple sheets
                processed_dfs = [extract_data_from_description(df) for df in data.values()]
                result_df = pd.concat(processed_dfs, ignore_index=True)
            else:
                result_df = extract_data_from_description(data)

            output_path = os.path.join(output_folder, file_name)
            result_df.to_excel(output_path, index=False)
            print(f"Processed: {file_name}")
        except Exception as e:
            print(f"Error processing {file_name}: {e}")

