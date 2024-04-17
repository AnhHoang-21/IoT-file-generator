import math
import pandas as pd
from tkinter import Tk, simpledialog

# Read the Excel file into a DataFrame
df = pd.read_excel("IoT keys.xlsx")

# Convert the DataFrame to a list of dictionaries
data = df.to_dict(orient='records')

# Function to replace nan values with 'nan'
def replace_nan_with_str(dictionary):
    for key, value in dictionary.items():
        if isinstance(value, float) and math.isnan(value):
            dictionary[key] = 'nan'
    return dictionary

# Transform each dictionary in the original list
transformed_data = [replace_nan_with_str(d) for d in data]

# Function to replace nan and 0.0 values with "nan" and "0" respectively
def replace_nan_and_zero_with_str(dictionary):
    for key, value in dictionary.items():
        if isinstance(value, float) and math.isnan(value):
            dictionary[key] = "nan"
        elif value == 0.0:
            dictionary[key] = "0"
    return dictionary

# Transform each dictionary in the original list
transformed_data = [replace_nan_and_zero_with_str(d) for d in data]

def construct_a_key(Protocol, Poll_Interval, Slave_IP_address, MBFC, Address, Number_of_register, Data_type, Multiplier, Adder, Key_Name, IOTMODE):
    chunk_code_of_key = 'TYPE,' + Protocol + ',' + str(Poll_Interval) + '\n' + 'ADDR,' + Slave_IP_address + '\n' + 'MBFC,' + str(MBFC) + '\n' + 'REGS,' + str(Address) + ',' + str(Number_of_register) + ',' + Data_type
    sort_NA_mutiplier= str(Multiplier)
    sort_NA_adder = str(Multiplier)

    # Sorting out the multiplier and adder
    if sort_NA_mutiplier == 'nan' and sort_NA_adder == 'nan':
        chunk_code_of_key =  chunk_code_of_key + '\n' + 'Key,' + Key_Name + '\n' + 'IOTMODE,' + str(IOTMODE)
    else:
        pass

    if sort_NA_mutiplier != 'nan':
        chunk_code_of_key = chunk_code_of_key + ',' + str(Multiplier)
    else:
        pass

    if sort_NA_adder != 'nan':
        chunk_code_of_key = chunk_code_of_key + ',' + str(Adder)  +  '\n' + 'Key,' + Key_Name + '\n' + 'IOTMODE,' + str(IOTMODE)
    else:
        pass
    return chunk_code_of_key

def get_slave_ip():
    """
    Uses a simple GUI to get the Slave IP address from the user.
    """
    root = Tk()
    root.withdraw()  # Hide the main window

    ip_address = simpledialog.askstring("Input", "Enter Slave IP address:")
    return ip_address

# Get the Slave IP address using GUI
Slave_IP_address = get_slave_ip()

# Open the file in write mode
file_name = "iotasset.txt"
with open(file_name, 'w') as f:
    # Iterate over each dictionary and call construct_a_key function
    for data in transformed_data:
        result = construct_a_key(data['Protocol'], data['Poll Inteval'], Slave_IP_address, data['MBFC'], data['Address'], data['  Number_of_register'], data[' Data_type'], data['Multiplier'], data['Adder'], data['Key_Name'], data['IOTMODE'])
        # Write the result to the text file
        f.write(result + '\n\n')  # Add '\n' between each result and an additional '\n' after each result

# Print confirmation message
print("Results have been saved to", file_name)
