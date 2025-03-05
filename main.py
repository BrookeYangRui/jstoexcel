import os
import json
import pandas as pd


def json_to_dataframe(json_file):
    """Convert a JSON file to a Pandas DataFrame with selected columns."""
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    df = pd.DataFrame(data)
    selected_columns = ['name', 'stargazers_count', 'html_url']
    df_filtered = df[selected_columns]

    df_filtered = df_filtered.rename(columns={
        'name': 'Application',
        'stargazers_count': 'Stars',
        'html_url': 'URL'
    })

    return df_filtered.drop_duplicates()


def process_json_files_in_directory(directory):
    """Process all JSON files in a directory and create individual and summary Excel files."""
    all_dataframes = []

    for filename in os.listdir(directory):
        if filename.endswith('.json'):
            json_path = os.path.join(directory, filename)
            df = json_to_dataframe(json_path)

            if not df.empty:
                excel_filename = os.path.splitext(filename)[0] + '.xlsx'
                excel_path = os.path.join(directory, excel_filename)
                df.to_excel(excel_path, index=False)
                print(f"Converted {filename} to {excel_filename}")

                all_dataframes.append(df)

    if all_dataframes:
        summary_df = pd.concat(all_dataframes, ignore_index=True).drop_duplicates()
        summary_excel_path = os.path.join(directory, 'Summary.xlsx')
        summary_df.to_excel(summary_excel_path, index=False)
        print(f"Summary Excel file created: {summary_excel_path}")
    else:
        print("No valid JSON files found.")


if __name__ == '__main__':
    directory = os.getcwd()  # Use the current working directory
    process_json_files_in_directory(directory)
