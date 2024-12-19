#pip install pandas matplotlib tk os datetime

import pandas as pd
from tkinter import Tk, filedialog, Listbox, MULTIPLE, Label, Checkbutton, IntVar, Button, Scrollbar, END, StringVar, simpledialog, DoubleVar
import os
from datetime import datetime
from datetime import datetime, timedelta
import openpyxl
from PIL import Image, ImageEnhance
import tifffile as tiff
import numpy as np

# Create the main Tkinter window
root = Tk()
root.title("Strain analysis")

# Variables~DCCC    
selected_folder = StringVar()
excel_file = StringVar()
echem_matched_path = StringVar()
con = DoubleVar()
shrp = DoubleVar()
final = StringVar()
discharge = IntVar()
charge = IntVar()

# Selects the folder containing the images, excel files, and Vic2D
def select_folder():
    strain_folder = filedialog.askdirectory(title="Select Strain folder")
    if strain_folder:
        selected_folder.set(strain_folder)
        charge.set(0)
        discharge.set(0)

# Selects the excel file containing the orgininal Arbin Data
def select_file():
    selected_file = filedialog.askopenfilename(title="Select excel file", filetypes=[("Excel files", "*.xlsx")])
    if selected_file:
        excel_file.set(selected_file)

def sort_data(selected_folder, excel_file):

    # Finds the first image in the folder to match the image time with the echem times
    images = [file for file in os.listdir(selected_folder.get()) if os.path.isfile(os.path.join(selected_folder.get(), file))]
    image_files = [file for file in images if os.path.splitext(file)[1].lower() == '.tiff']
    if image_files:
        image1_path = os.path.join(selected_folder.get(), image_files[0])
    else:
        image1_path = None  # Assign a default value if no image files are found
    
    if image1_path:
    # Extract the date and time from the image name
        image1_date_time_str = image1_path.split('__')[1].split('.')[0].replace('_', '-')
        image_date_time = datetime.strptime(image1_date_time_str, '%Y-%m-%d-%H-%M-%S')

    # Reads and extracts relavent information from the excel file
    channel_num = pd.read_excel(excel_file.get(), sheet_name = 'Global_Info').iloc[3,0]
    df = pd.read_excel(excel_file.get(), sheet_name=f'Channel_{channel_num}_1')
    mass = pd.read_excel(excel_file.get(), sheet_name = 0).iloc[3,7]

    # Adjusts the 'Date_Time' column to match the time from the image names and adds a data point column
    if not 'data_pt' in df.columns:
        df.insert(0, 'data_pt', range(1, len(df) + 1))
    df['Date_Time'] = pd.to_datetime(df['Date_Time'])
    first_cell_value = pd.to_datetime(df.iloc[0, 1])
    time_adjustment = first_cell_value - image_date_time
    df.iloc[:, 1] = pd.to_datetime(df.iloc[:, 1]) - time_adjustment

    # Sets OCV step to cycle 0
    df.loc[(df['Step_Index'] == 1) & (df['Cycle_Index'] == 1), 'Cycle_Index'] = 0

    # Create a DataFrame to store matched rows
    matched_rows = pd.DataFrame()

    # Define the path for the matched echem data to be exported
    output_excel_file = os.path.join(selected_folder.get(), 'echem data matched.xlsx')

    # Iterate over the images in the folder
    for image_name in os.listdir(selected_folder.get()):
        if image_name.endswith('.tiff'):
            # Extract the date and time from the image name
            date_time_str = image_name.split('__')[1].split('.')[0].replace('_', '-')
            image_date_time = datetime.strptime(date_time_str, '%Y-%m-%d-%H-%M-%S')

            # Find the matching row in the Excel file
            df['Date_Time'] = pd.to_datetime(df['Date_Time'])
            df['time_diff'] = (df['Date_Time'] - image_date_time).abs()
            # Get the row with the minimum time difference
            matching_row = df.loc[df['time_diff'].idxmin()]
            # Check if the time difference is greater than 300 seconds
            if matching_row['time_diff'] > timedelta(seconds=305):
                matching_row = pd.DataFrame()  
            
            if not matching_row.empty:
                # Append the matching row to the matched_rows DataFrame
                matched_rows = pd.concat([matched_rows, matching_row.to_frame().T])
            
            # If image was taken outside of the testing parameters, remove the image from the data set
            if matching_row.empty:
                subfolder_name = 'Extra images'
                subfolder_path = os.path.join(selected_folder.get(), subfolder_name)
                if not os.path.exists(subfolder_path):
                    os.makedirs(subfolder_path)
                    # Move the image to the subfolder
                old_path = os.path.join(selected_folder.get(), image_name)
                new_path = os.path.join(subfolder_path, image_name)
                os.rename(old_path, new_path)
    
    # Adjust the data based on the mass
    echem_data_sorted = matched_rows
    unit = 'g'
    echem_data_sorted['Test Time(h)'] = (echem_data_sorted['Test_Time(s)'] / 3600)
    echem_data_sorted[f'Current (mA/{unit})'] = (echem_data_sorted['Current(A)'] * 1000) / mass
    echem_data_sorted[f'Power (mW/{unit})'] = (echem_data_sorted['Current(A)'] * echem_data_sorted['Voltage(V)'] * 1000) / mass
    echem_data_sorted[f'Charge Capacity (mAh/{unit})'] = (echem_data_sorted['Charge_Capacity(Ah)'] * 1000) / mass
    echem_data_sorted[f'Discharge Capacity (mAh/{unit})'] = (echem_data_sorted['Discharge_Capacity(Ah)'] * 1000) / mass
    echem_data_sorted[f'Charge Energy (mWh/{unit})'] = (echem_data_sorted['Charge_Energy(Wh)'] * 1000) / mass
    echem_data_sorted[f'Discharge Energy (mWh/{unit})'] = (echem_data_sorted['Discharge_Energy(Wh)'] * 1000) / mass
    echem_data_sorted['Ref. Voltage (V)'] = (echem_data_sorted.iloc[:,15])
    # Drop the original columns
    echem_data_sorted.drop(columns=['Test_Time(s)', 'time_diff', 'Current(A)', 'Charge_Capacity(Ah)', 'Discharge_Capacity(Ah)', 'Charge_Energy(Wh)', 'Discharge_Energy(Wh)', 'Internal Resistance(Ohm)', 'dV/dt(V/s)','BuiltIn_Aux_Voltage1(V)', 'Unnamed: 14'], inplace=True)
    echem_data_sorted.insert(2, 'Test Time(h)', echem_data_sorted.pop('Test Time(h)'))

    # Save the matched rows to a new Excel file
    echem_data_sorted.to_excel(output_excel_file, index=False)
    echem_matched_path.set(output_excel_file)

    # Save edits to the original Excel file
    with pd.ExcelWriter(excel_file.get(), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=f'Channel_{channel_num}_1', index=False)
    print('Data sorted successfully')

# Boosts the contrast and sharpness of images to improve the speckle pattern
def contrast(selected_folder):
    # Prompt to enter contrast and sharpness
    con = simpledialog.askfloat("Input", "Enter contrast boost value:")
    shrp = simpledialog.askfloat("Input", "Enter sharpness boost value:")
    bright = simpledialog.askfloat("Input", "Enter Brightness boost value:")

    # Loops through all images in the strain folder and sub folders
    for root, dirs, files in os.walk(selected_folder.get()):
        for filename in files:
            if filename.endswith(('.tiff')):
                # Open an image file
                with tiff.TiffFile(os.path.join(root, filename)) as tif:
                    img = tif.asarray()
                 
                # Convert to PIL Image for processing
                img = Image.fromarray(img)

                # Enhance contrast
                enhancer = ImageEnhance.Contrast(img)
                img = enhancer.enhance(con)  # Adjust the contrast value as needed

                # Enhance sharpness
                enhancer = ImageEnhance.Sharpness(img)
                img = enhancer.enhance(shrp)  # Adjust the sharpness value as needed

                # Enhance sharpness
                enhancer = ImageEnhance.Brightness(img)
                img = enhancer.enhance(bright)  # Adjust the brightness value as needed

                # Save the modified image to the original location
                img.save(os.path.join(root, filename), format='TIFF')
    print('Contrasst boosted successfully')

# Sorts images and VIC 2D files into cycle folders
def sort_images(selected_folder, echem_matched_path):
    # Loads the matched/sorted echem data if not done already
    echem_sorted_path = os.path.join(selected_folder.get(), 'echem data matched.xlsx')
    echem_matched_path.set(echem_sorted_path)
    echem_data_matched = pd.read_excel(echem_sorted_path)
    
    # Iterate over the images in the folder
    for name in os.listdir(selected_folder.get()):
        if name.endswith(('.tiff', '.csv')):
            # Extract the date and time from the image name
            date_time_str = name.split('__')[1].split('.')[0].replace('_', '-')
            image_date_time = datetime.strptime(date_time_str, '%Y-%m-%d-%H-%M-%S')

            # Find the matching row in the Excel file
            echem_data_matched['Date_Time'] = pd.to_datetime(echem_data_matched['Date_Time'])
            echem_data_matched['time_diff'] = (echem_data_matched['Date_Time'] - image_date_time).abs()
            # Get the row with the minimum time difference
            matching_row = echem_data_matched.loc[echem_data_matched['time_diff'].idxmin()]

            if not matching_row.empty:
                # Get the Cycle_Index value
                cycle_index = matching_row['Cycle_Index']

                # Create the subfolder if it doesn't exist
                subfolder_name = 'OCV' if cycle_index == 0 else f'Cycle {cycle_index}'
                subfolder_path = os.path.join(selected_folder.get(), subfolder_name)
                os.makedirs(subfolder_path, exist_ok=True)
                
                # Create the image folder if it doesn't exist
                imagesubfolder_path = os.path.join(subfolder_path, 'Images')
                os.makedirs(imagesubfolder_path, exist_ok=True)

                # Create the DIC data folder if it doesn't exist
                DICsubfolder_path = os.path.join(subfolder_path, 'VIC 2D export data')
                os.makedirs(DICsubfolder_path, exist_ok=True)

                # Move the file to the appropriate subfolder
                old_path = os.path.join(selected_folder.get(), name)
                new_path = os.path.join(imagesubfolder_path, name) if name.endswith('.tiff') else os.path.join(DICsubfolder_path, name)
                os.rename(old_path, new_path)
    print('Cycles divided successfully')


# Function to safely pad a list with a specific value
def pad_column(col, max_length, pad_value):
    # Convert to NumPy array before padding
    col = np.array(col)
    return np.pad(col, (0, max_length - len(col)), 'constant', constant_values=pad_value)

# Function to convert columns to appropriate types safely
def convert_to_appropriate_type_safely(col):
    if isinstance(col, list):
        col = pd.Series(col)
    if pd.api.types.is_numeric_dtype(col):
        return col.fillna(-1 if pd.api.types.is_integer_dtype(col) else np.nan).astype(int if pd.api.types.is_integer_dtype(col) else float)
    else:
        return col
     

def post_analysis(selected_folder, echem_matched_path, charge, discharge):
    # Load the matched/sorted echem data if not done already
    echem_sorted_path = os.path.join(selected_folder.get(), 'echem data matched.xlsx')
    echem_matched_path.set(echem_sorted_path)
    echem_data_matched = pd.read_excel(echem_sorted_path)

    # Determine the number of cycles in the test
    num_cycles = echem_data_matched['Cycle_Index'].max()

    # Initialize dictionaries to store DataFrames
    df_chg_dict = {}
    df_disc_dict = {}
    df_disc_test = {}
    df_chg_test = {}
    data = {}
    df_Test_data = []
    df_OCV = []
    df_Time = []
    df_exx = []
    df_exy = []
    df_eyy = []
    df_test_time = []
    df_step_time = []
    df_data_pt = []
    df_Date_Time = []
    df_Volt = []
    df_current = []
    df_cap = []
    df_cycle = []
    df_step = []

    # Loop through all files in the strain folder
    for root, dirs, files in os.walk(selected_folder.get()):
        for filename in files:
            if filename.endswith('.csv'):
                # Access each data point
                df = pd.read_csv(os.path.join(root, filename))

                # Remove extra spaces and quotation marks from column names
                df.columns = df.columns.str.strip().str.replace('"', '')

                # Remove uncorrelated points
                df.loc[df['sigma'] == -1, ['exx', 'exy', 'eyy']] = np.nan

                # Average the strain data
                exx_data = np.nanmean(df['exx']) * 100
                exy_data = np.nanmean(df['exy']) * 100
                eyy_data = np.nanmean(df['eyy']) * 100

                # Pull the time of the data point from the file name
                date_time_str = filename.split('__')[1].split('.')[0].replace('_', '-')
                strain_date_time = datetime.strptime(date_time_str, '%Y-%m-%d-%H-%M-%S')

                # Append the strain data
                df_Time.append(strain_date_time)
                df_exx.append(exx_data)
                df_exy.append(exy_data)
                df_eyy.append(eyy_data)

                # Match the strain and the echem data
                echem_data_matched['Date_Time'] = pd.to_datetime(echem_data_matched['Date_Time'])
                echem_data_matched['time_diff'] = (echem_data_matched['Date_Time'] - strain_date_time).abs()
                # Get the row with the minimum time difference
                matching_row = echem_data_matched.loc[echem_data_matched['time_diff'].idxmin()]

                # Append the echem data
                df_data_pt.append(matching_row['data_pt'])
                df_Date_Time.append(matching_row['Date_Time'])
                df_test_time.append(matching_row['Test Time(h)'])
                df_step_time.append(matching_row['Step_Time(s)'])
                df_cycle.append(matching_row['Cycle_Index'])
                df_step.append(matching_row['Step_Index'])
                df_Volt.append(matching_row['Voltage(V)'])
                df_current.append(matching_row['Current (mA/g)'])

                if matching_row['Step_Index'] == 1:
                    df_cap.append(matching_row['Discharge Capacity (mAh/g)'])
                elif discharge.get() == 1:
                    if matching_row['Step_Index'] == 2:
                        df_cap.append(matching_row['Discharge Capacity (mAh/g)'])
                    if matching_row['Step_Index'] == 3:
                        df_cap.append(matching_row['Charge Capacity (mAh/g)'])
                elif charge.get() == 1:
                    if matching_row['Step_Index'] == 3:
                        df_cap.append(matching_row['Discharge Capacity (mAh/g)'])
                    if matching_row['Step_Index'] == 2:
                        df_cap.append(matching_row['Charge Capacity (mAh/g)'])

    # Define all your column lists
    columns = [
        df_data_pt, df_Date_Time, df_test_time, df_step_time, df_cycle, df_step, df_Volt, df_current, df_cap, df_exx, df_eyy, df_exy
    ]

    # Convert all columns to appropriate types safely before padding
    converted_columns = [convert_to_appropriate_type_safely(col) for col in columns]

    # Calculate max length
    max_length = max(len(col) for col in converted_columns)

    # Pad all columns to max length with appropriate value for each type
    padded_columns = [pad_column(col, max_length, -1000 if pd.api.types.is_integer_dtype(col) else np.nan) for col in converted_columns]

    # Create a DataFrame for the padded data
    data['Test'] = pd.DataFrame({
        'Data Point': padded_columns[0],
        'Date Time' : padded_columns[1],
        'Test Time (h)': padded_columns[2],
        'Step Time (s)': padded_columns[3],
        'Cycle Index': padded_columns[4],
        'Step Index': padded_columns[5],
        'Voltage (V)': padded_columns[6],
        'Current Density (mA/g)': padded_columns[7],
        'Specific Capacity (mAh/g)': padded_columns[8],
        'exx (%)': padded_columns[9],
        'eyy (%)': padded_columns[10],
        'exy (%)': padded_columns[11],
    })

    # Add the OCV data
    OCV_data = data['Test'][data['Test']['Cycle Index'] == 0]

    # Loops through the cycles
    for i in range(1, num_cycles + 1):
        # Compiles data if the discharge cycle comes first
        if discharge.get() == 1:
            # Reference the echem data
            cycle_data = data['Test'][data['Test']['Cycle Index'] == i]
            echem_disc_data = cycle_data[cycle_data['Step Index'] == 2]
            echem_chg_data = cycle_data[cycle_data['Step Index'] == 3]

            # Create DataFrames for data
            df_disc_dict[f'disc_{i}'] = echem_disc_data
            df_chg_dict[f'chg_{i}'] = echem_chg_data

            # References the first row in each data frame
            first_valid = df_disc_dict[f'disc_{i}'].first_valid_index()
            
            # If not the first cycle, add the final value of the previous cycle to the current cycle if the data is not continuous
            if i > 1 and not df_disc_dict[f'disc_{i}'].empty and df_disc_dict[f'disc_{i}'].loc[first_valid, 'exx (%)'] == 0:
                for col in ['exx (%)', 'eyy (%)', 'exy (%)']:
                    last_valid = df_chg_dict[f'chg_{i-1}'][col].last_valid_index()
                    if f'disc_{i}' not in df_disc_test:
                        df_disc_test[f'disc_{i}'] = df_disc_dict[f'disc_{i}'].copy()
                    if f'chg_{i}' not in df_chg_test:
                        df_chg_test[f'chg_{i}'] = df_chg_dict[f'chg_{i}'].copy()
                    if last_valid is not None:
                        df_disc_test[f'disc_{i}'].loc[:, col] += df_chg_test[f'chg_{i-1}'][col].loc[last_valid]
                        df_chg_test[f'chg_{i}'].loc[:, col] += df_chg_test[f'chg_{i-1}'][col].loc[last_valid]

            # Reset strain to 0 on each new cycle if analysis was continuous
            elif first_valid is not None and df_disc_dict[f'disc_{i}'].loc[first_valid, 'exx (%)'] != 0:
                df_disc_test[f'disc_{i}'] = df_disc_dict[f'disc_{i}'].copy()
                df_chg_test[f'chg_{i}'] = df_chg_dict[f'chg_{i}'].copy()
                for col in ['exx (%)', 'eyy (%)', 'exy (%)']:
                    reset = df_disc_dict[f'disc_{i}'].loc[first_valid, col]
                    df_disc_dict[f'disc_{i}'].loc[:, col] -= reset
                    df_chg_dict[f'chg_{i}'].loc[:, col] -= reset
            else:
                df_disc_test[f'disc_{i}'] = df_disc_dict[f'disc_{i}'].copy()
                df_chg_test[f'chg_{i}'] = df_chg_dict[f'chg_{i}'].copy()

            # Append the DataFrames to the list
            df_Test_data.append(df_disc_test[f'disc_{i}'])
            df_Test_data.append(df_chg_test[f'chg_{i}'])

        # Compiles data if the charge cycle comes first
        if charge.get() == 1:
            # Reference the echem data
            cycle_data = data['Test'][data['Test']['Cycle Index'] == i]
            echem_disc_data = cycle_data[cycle_data['Step Index'].isin([3, 6, 9])]
            echem_chg_data = cycle_data[cycle_data['Step Index'].isin([2, 5, 8])]
            #relax_data = data['Test'][data['Test']['Cycle Index'] == 11]

            # Create DataFrames for data
            df_disc_dict[f'disc_{i}'] = echem_disc_data
            df_chg_dict[f'chg_{i}'] = echem_chg_data

            # References the first row in each data frame
            first_valid = df_chg_dict[f'chg_{i}'].first_valid_index()
            
            # If not the first cycle, add the final value of the previous cycle to the current cycle if the data is not continuous
            if i > 1 and not df_chg_dict[f'chg_{i}'].empty and df_chg_dict[f'chg_{i}'].loc[first_valid, 'exx (%)'] == 0:
                for col in ['exx (%)', 'eyy (%)', 'exy (%)']:
                    last_valid = df_disc_dict[f'disc_{i-1}'][col].last_valid_index()
                    if f'chg_{i}' not in df_chg_test:
                        df_chg_test[f'chg_{i}'] = df_chg_dict[f'chg_{i}'].copy()
                    if f'disc_{i}' not in df_disc_test:
                        df_disc_test[f'disc_{i}'] = df_disc_dict[f'disc_{i}'].copy()
                    if last_valid is not None:
                        df_chg_test[f'chg_{i}'].loc[:, col] += df_disc_test[f'disc_{i-1}'][col].loc[last_valid]
                        df_disc_test[f'disc_{i}'].loc[:, col] += df_disc_test[f'disc_{i-1}'][col].loc[last_valid]

            # Reset strain to 0 on each new cycle if analysis was continuous
            elif first_valid is not None and df_chg_dict[f'chg_{i}'].loc[first_valid, 'exx (%)'] != 0:
                df_chg_test[f'chg_{i}'] = df_chg_dict[f'chg_{i}'].copy()
                df_disc_test[f'disc_{i}'] = df_disc_dict[f'disc_{i}'].copy()
                for col in ['exx (%)', 'eyy (%)', 'exy (%)']:
                    reset = df_chg_dict[f'chg_{i}'].loc[first_valid, col]
                    df_chg_dict[f'chg_{i}'].loc[:, col] -= reset
                    df_disc_dict[f'disc_{i}'].loc[:, col] -= reset
            else:
                df_chg_test[f'chg_{i}'] = df_chg_dict[f'chg_{i}'].copy()
                df_disc_test[f'disc_{i}'] = df_disc_dict[f'disc_{i}'].copy()

            # Append the DataFrames to the list
            df_Test_data.append(df_chg_test[f'chg_{i}'])
            df_Test_data.append(df_disc_test[f'disc_{i}'])


    # Concatenate the list of DataFrames into a single DataFrame
    df_OCV.append(OCV_data)
    test_data = pd.concat(df_OCV + df_Test_data, ignore_index=True)

    # Remove blank rows
    Test_data = test_data.dropna(how='all')

    # Removes OCV expansion from full test data
    first_valid_test = Test_data.first_valid_index()
    if first_valid_test is not None and Test_data.loc[first_valid_test, 'exx (%)'] != 0:
        for col in ['exx (%)', 'eyy (%)', 'exy (%)']:
            reset_test = Test_data.loc[first_valid_test, col]
            Test_data.loc[:, col] -= reset_test

    # Define the path for the Excel file in the Test_directory
    excel_file_path = os.path.join(selected_folder.get(), 'Echem_and_Strain_Combined.xlsx')

    # Write all DataFrames to separate sheets in the Excel file
    with pd.ExcelWriter(excel_file_path) as writer:
        OCV_data.to_excel(writer, sheet_name='OCV', index=False)
        for key, df in df_disc_dict.items():
            df.to_excel(writer, sheet_name=key, index=False)
        for key, df in df_chg_dict.items():
            df.to_excel(writer, sheet_name=key, index=False)
        Test_data.to_excel(writer, sheet_name='Full Test', index=False)
    print('Analysis Finished')

# Define button commands
def sort():
    sort_data(selected_folder, excel_file)

def boost():
    contrast(selected_folder)

def divide():
    sort_images(selected_folder, echem_matched_path)

def finish():
    post_analysis(selected_folder, echem_matched_path, charge, discharge)
 
# Add all GUI elements to the main window
Label(root, text="Select a strain folder to begin:").grid(row=0, column=0, columnspan=2)
Button(root, text="Select strain folder", command=select_folder).grid(row=1, column=0, columnspan=2)
Label(root, textvariable=selected_folder).grid(row=2, column=0, pady=10, columnspan=2)

Label(root, text="Select the Arbin Export file to begin:").grid(row=3, column=0, columnspan=2)
Button(root, text="Select excel file", command=select_file).grid(row=4, column=0, columnspan=2)
Label(root, textvariable=excel_file).grid(row=5, column=0, pady=10, columnspan=2)

Label(root, text="Select whether the charge or discharge step occurs first :").grid(row=6, column=0, columnspan=2)
Checkbutton(root, text="Charge", variable=charge).grid(row=7, column=0)
Checkbutton(root, text="Discharge", variable=discharge).grid(row=7, column=1, pady=10)

Label(root, text="Sort the echem data to include only the data relavent to images:").grid(row=8, column=0, columnspan=2)
Button(root, text="Sort Echem data", command=sort).grid(row=9, column=0, pady=10, columnspan=2)

Label(root, text="Boosts the contrast and sharpness of the images:").grid(row=10, column=0, columnspan=2)
Label(root, text= "this may take some time:").grid(row=11,column=0, columnspan=2)
Button(root, text="Boost contrast", command=boost).grid(row=12, column=0, columnspan=2)
Label(root, textvariable=f'Contrast and Sharpness inc by factors of {con} and {shrp}').grid(row=13, column=0, pady=10)

Label(root, text="Sorts images and strain data files into cycle folders:").grid(row=14, column=0, columnspan=2)
Button(root, text="Sort into cycle folders", command=divide).grid(row=15, column=0, pady=10, columnspan=2)

Label(root, text="Convert VIC 2D and echem data into a single Excel File:").grid(row=16, column=0, columnspan=2)
Button(root, text="Finish Analysis", command=finish).grid(row=17, column=0, columnspan=2)
Label(root, textvariable=final).grid(row=18, column=0, columnspan=2)

root.mainloop()
