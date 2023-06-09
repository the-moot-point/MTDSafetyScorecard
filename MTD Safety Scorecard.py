from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import datetime
import os
import glob
from dateutil.relativedelta import relativedelta
from pathlib import Path


def convert_to_seconds(time_string):
    h, m, s = map(int, time_string.split(':'))
    return h * 3600 + m * 60 + s


def assign_coaching_videos(df):
    # Initialize a new column in the existing DataFrame to hold the results
    df['Formal Warning Video Assignment'] = ''

    # Criteria
    criteria = [
        # Criterion 1
        ((df['Percent Moderate Speeding'] +
          df['Percent Heavy Speeding'] +
          df['Percent Severe Speeding'] > 9.9), 'Speeding'),
        # Criterion 2
        ((df['Mobile Usage'] >= 5), 'Werner Herzog Movie'),
        # Criterion 3
        ((df['Inattentive Driving'] >= 5), 'End Distracted Driving'),
        # Criterion 4
        ((df[['Did Not Yield (Manual)',
              'Ran Red Light (Manual)',
              'Rolling Stop',
              'Lane Departure (Manual)']].sum(axis=1) >= 5), 'Rolling Stops'),
        # Criterion 5
        ((df[['Following of 0-2s (Manual)', 'Following of 2-4s (Manual)', 'Late Response (Manual)',
              'Defensive Driving (Manual)', 'Following Distance']].sum(axis=1) >= 5),
         'Tailgating and the 3-second Rule'),
        # Criterion 6
        ((df['No Seat Belt'] >= 3), 'Seatbelt Use')
    ]

    # Apply each criterion
    for criterion, video in criteria:
        df.loc[criterion, 'Formal Warning Video Assignment'] += video + ', '

    # Remove trailing comma and space from assignments
    df['Formal Warning Video Assignment'] = df['Formal Warning Video Assignment'].str.rstrip(', ')

    return df


def get_unique_filename(filepath: str) -> str:
    filename, extension = os.path.splitext(filepath)
    counter = 1

    while os.path.exists(filepath):
        filepath = f"{filename}({counter}){extension}"
        counter += 1

    return filepath


def read_config(file_path: Path):
    """
    Function to read the configuration file and extract constants.
    """
    constants = {}

    with open(file_path, 'r') as file:
        for line in file:
            key, value = line.strip().split('=')
            constants[key] = int(value)

    return constants


def read_data(file_path):
    df = pd.read_excel(file_path)
    return df


def split_driver_tags(df):
    df[['Company', 'Location', 'Peer Group']] = df['Driver Tags'].str.split(',', expand=True)
    df[['Company', 'Location', 'Peer Group']] = df[['Company', 'Location', 'Peer Group']].apply(lambda x: x.str.strip())
    return df


def create_filtered_report(df, tag, columns):
    """
    Function to create a report for drivers with a specific tag.
    """
    # Filter the DataFrame based on the tag
    filtered_df = df[df['Driver Tags'].str.contains(tag, na=False)]

    # Create a dictionary for the report
    report_dict = {}

    # Select columns for the report
    for column in columns:
        report_dict[column] = filtered_df[column]

    # Create the report DataFrame
    filtered_report = pd.DataFrame(report_dict)

    return filtered_report


def export_to_excel(driver_safety_report, output_path):
    with pd.ExcelWriter(output_path) as writer:
        driver_safety_report.to_excel(writer, sheet_name='Driver Safety', index=False)


def get_latest_file_in_directory(directory, *extensions):
    # Find all files in the directory with any of the specified extensions
    files = []
    for extension in extensions:
        files.extend(glob.glob(f"{directory}/*.{extension}"))

    # Get the newest file (based on modification time)
    newest_file = max(files, key=os.path.getmtime)

    return newest_file


def create_warning_report(df, columns):
    """
    Function to create a formal warning report.
    Includes drivers with scores 70 or below or with a recorded crash.
    """
    warning_criteria = (df['Safety Score'] <= 70) | (df['Crash'] > 0)
    warning_report = df.loc[warning_criteria, columns]

    return warning_report


def create_perfect_scores(df, columns):
    """
    Function to create the perfect scores report.
    Includes valid drivers with 100 safety score.
    """
    valid_peers = ['Driver', 'Reset', 'Warehouse']
    perfect_criteria = (df['Peer Group'].isin(valid_peers)) & (df['Safety Score'] >= 100)
    perfect_report = df.loc[perfect_criteria, columns]

    return perfect_report

def score_range(score):
    if score == 100:
        return "Perfect 100"
    elif score > 70:
        return "Above 70"
    elif 36 <= score <= 70:
        return "Below 70"
    elif score <= 35:
        return "Critical - Below 35"


def main():
    """
    Main function to read the data and generate driver safety report.
    """
    # Specify the directory where the .xlsx files are stored
    directory = 'C:/Users/sgtjo/Documents/Samsara MTD Scorecard/Samsara _raw_data'

    # Get the newest .xlsx file in the directory
    input_file_path = get_latest_file_in_directory(directory, 'xlsx')

    # Read the config file
    directory_path = Path(directory)
    config = read_config(directory_path.parent / 'config.txt')

    # Read the data
    df = read_data(input_file_path)

    # Split the driver tags
    df = split_driver_tags(df)

    # Add new column of drive times in seconds
    df['Drive Time (seconds)'] = df['Drive Time (hh:mm:ss)'].apply(convert_to_seconds)

    # Filter out drivers with less than the configured minimum drive time
    df = df[df['Drive Time (seconds)'] >= config['MIN_DRIVE_TIME']]

    # Make the summed up columns
    df['Collision Risk'] = df['Following Distance'] + df['Late Response (Manual)'] + df['Near Collision (Manual)']
    df['Harsh Events'] = df['Harsh Accel'] + df['Harsh Brake'] + df['Harsh Turn']
    df['Traffic Violations'] = df['Rolling Stop'] + df['Did Not Yield (Manual)'] + df['Ran Red Light (Manual)'] + \
                               df['Lane Departure (Manual)']
    df['Policy Violations'] = df['Obstructed Camera (Automatic)'] + df['Obstructed Camera (Manual)'] + df[
        'Eating/Drinking (Manual)'] + df['Smoking (Manual)'] + df['No Seat Belt']
    df['Speeding %'] = df['Percent Moderate Speeding'] + df['Percent Heavy Speeding'] + df['Percent Severe Speeding']
    df['Score Range'] = df['Safety Score'].apply(score_range)

    # Calculate formal warning video assignments
    df = assign_coaching_videos(df)

    # Create a filtered DataFrame for each report
    driver_scorecard = df[df['Driver Tags'].str.contains("Driver|Reset|Warehouse", na=False)].copy()
    manager_scorecard = df[df['Driver Tags'].str.contains("Manager", na=False)].copy()
    warning_report = df[(df['Safety Score'] <= 70) | (df['Crash'] > 0)].copy()
    coaching_video = df[df['Formal Warning Video Assignment'].str.strip() != '']
    perfect_report = df[
        (df['Peer Group'].str.contains("Driver|Reset|Warehouse")) & (df['Safety Score'] == 100)
        ].copy()
    top_drivers = df.sort_values(by='Mobile Usage', ascending=False).head(10)
    total_mobile_usage = df['Mobile Usage'].sum()
    top_drivers['Percent of Total Mobile Usage'] = top_drivers['Mobile Usage'] / total_mobile_usage * 100

    # Define columns for the driver scorecard dataframe
    scorecard_columns = ['Score Range', 'Company', 'Location', 'Driver Name', 'Peer Group',
                         'Safety Score', 'Drive Time (hh:mm:ss)', 'Percent Moderate Speeding',
                         'Percent Heavy Speeding', 'Percent Severe Speeding',
                         'Mobile Usage', 'Crash', 'Collision Risk', 'Harsh Events',
                         'Inattentive Driving', 'Traffic Violations', 'Policy Violations']

    coaching_columns = ['Company', 'Location', 'Driver Name', 'Peer Group',
                        'Safety Score', 'Drive Time (hh:mm:ss)', 'Speeding %',
                        'Mobile Usage', 'Crash', 'Collision Risk', 'Harsh Events',
                        'Inattentive Driving', 'Traffic Violations', 'Policy Violations',
                        'Formal Warning Video Assignment']

    top_10_columns = ['Driver Name', 'Mobile Usage', 'Percent of Total Mobile Usage']

    perfect_columns = ['Company', 'Location', 'Driver Name', 'Safety Score', 'Drive Time (hh:mm:ss)']

    driver_scorecard = driver_scorecard.loc[:, scorecard_columns]
    manager_scorecard = manager_scorecard.loc[:, scorecard_columns]
    warning_report = warning_report.loc[:, scorecard_columns]
    coaching_video_df = coaching_video.loc[:, coaching_columns]
    perfect_report = perfect_report.loc[:, perfect_columns]
    top_drivers = top_drivers.loc[:, top_10_columns]

    # Prepare the reports list with titles and dataframes
    reports = [
        {'title': 'Driver Scorecard', 'dataframe': driver_scorecard},
        {'title': 'Manager Scorecard', 'dataframe': manager_scorecard},
        {'title': 'Paycom - Formal Warning', 'dataframe': warning_report},
        {'title': 'Paycom - Coaching Videos', 'dataframe': coaching_video_df},
        {'title': 'Perfect Scores', 'dataframe': perfect_report},
        {'title': 'Top 10 Mobile Usage', 'dataframe': top_drivers}

    ]

    # Define the workbook in memory
    wb = Workbook()

    # remove the default sheet created and keep our sheets only
    wb.remove(wb.active)

    # create reports (sheets) and add them to the workbook
    for report in reports:
        ws = wb.create_sheet(title=report['title'])

        # add data to the sheet
        for r in dataframe_to_rows(report['dataframe'], index=False, header=True):
            ws.append(r)

    # Get the previous month
    current_month = datetime.datetime.now() - relativedelta(months=0)

    # Format the output file path
    directory_path = Path(directory)
    output_file_path = directory_path.parent / 'Samsara MTD Scorecard' / f"MTD Safety Scorecard - {current_month.strftime('%b %Y')}.xlsx"
    print(f"Output file path: {output_file_path}")

    # If file already exists, append a number suffix
    directory_path = Path(directory)
    if output_file_path.is_file():
        counter = 1
        while output_file_path.is_file():
            output_file_path = directory_path.parent / 'Samsara MTD Scorecard' / f" MTD Safety Scorecard - {current_month.strftime('%b %Y')} ({counter}).xlsx"
            counter += 1

    wb.save(output_file_path)


if __name__ == "__main__":
    main()

