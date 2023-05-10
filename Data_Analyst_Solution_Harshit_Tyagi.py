import os
import pandas as pd
import re
import csv
import argparse


class TooManyUniqueValuesException(BaseException):
    pass


# Function that returns a list for particular column from the merged excel data
def get_coulmn_data(dfs, sheet_name, column_number):
    column_data = []
    for df in dfs:
        sheet = df[sheet_name]
        column = sheet.iloc[3:, column_number]
        column = column.tolist()
        column_data.extend(column)
    return column_data



# Function to extract the result, Drawing_Index, Drawing_Type, Dp_Number
# Directory here top level directory like Res0014_AF_0_RTZ_001.pdf-Results
def extract_drawing_info(directory):
    directory = directory.split('.')[0]
    pattern = r"Res\d+_(\w+?)_(\d+?)_(\w+?)_\d+"

    match = re.match(pattern, directory)

    if match:
        Drawing_Index = match.group(1)
        Drawing_Type = match.group(3)
        Dp_Number = directory[directory.rfind("_") + 1:]
    else:
        Drawing_Index = ""
        Dp_Number = ""
        Drawing_Type = ""

    return directory, Drawing_Index, Drawing_Type, Dp_Number


# Function to extract the drawing_id
def get_drawing_id(df):
    drawing_id = get_coulmn_data(df, 'TitleBlocks', 5)
    # first_drawing_id = next(filter(lambda x: str(x) != 'nan', drawing_id),'')
    for id in drawing_id:
        if str(id) != 'nan':
            return repr(id)
            break

    return


# Function to extract the weight
def get_weight(df):
    weight = get_coulmn_data(df, 'TitleBlocks', 16)
    weight = set(weight)
    unique_weight = ""
    isUnique = False
    if (len(weight) > 2):
        raise TooManyUniqueValuesException("More than 1 unique value found")

    for weight in weight:
        if (str(weight) == 'nan' and not isUnique):
            pass
        else:
            unique_weight = weight
            isUnique = True
    if (str(unique_weight).isdigit()):
        return unique_weight

    return unique_weight


# Function to get the SHAPE counts and Tolerance
def get_shape_count(df):
    shape_column = get_coulmn_data(df, 'Measures', 5)
    tolerance_grade = get_coulmn_data(df, 'Measures', 12)

    upper_derivation_column = get_coulmn_data(df, 'Measures', 9)

    lower_derivation_column = get_coulmn_data(df, 'Measures', 10)

    count_dict = {'Dimension_Type_' + x + '_Count': shape_column.count(x) for x in set(shape_column)}



## Following IF conditions are written to print 0, if a particular shape is not present in the excels.
## If instead of 0, blank column is accepted, we can comment or remove the below IF conditions. There will be no affect on the code.
    if('Dimension_Type_ROUND_Count' not in count_dict):
        count_dict['Dimension_Type_ROUND_Count']=0

    if('Dimension_Type_NOMINAL_Count' not in count_dict):
        count_dict['Dimension_Type_NOMINAL_Count']=0

    tolerance_dict = {}
    for shape in shape_column:
        # Check for values in column M
        it_values = []
        for i in range(len(shape_column)):
            if shape_column[i] == shape and isinstance(tolerance_grade[i], str) and tolerance_grade[i].startswith('IT'):
                # Check if the string following "IT" is a float
                try:
                    it_value = int(tolerance_grade[i][2:])
                    it_values.append(it_value)
                except ValueError:
                    pass
        if len(it_values) > 0:
            # Take the smallest IT value as the tolerance
            tolerance = min(it_values)
        else:
            # Check values in columns J and K
            deviation_values = []
            for i in range(len(shape_column)):
                if shape_column[i] == shape:
                    upper_derivation = upper_derivation_column[i]
                    lower_derivation = lower_derivation_column[i]
                    if(str(upper_derivation) != 'nan' and str(lower_derivation) != 'nan'):
                        deviation = float(upper_derivation) - float(lower_derivation)
                        deviation_values.append(deviation)
                    elif(str(upper_derivation)=='nan' and str(lower_derivation) != 'nan' ):
                        deviation = 0-float(lower_derivation)
                        deviation_values.append(deviation)
                    elif (str(upper_derivation) != 'nan' and str(lower_derivation) == 'nan'):
                        deviation = float(upper_derivation) - 0
                        deviation_values.append(deviation)

                    # deviation_values.append(deviation)
            if len(deviation_values) > 0:
                # Take the smallest deviation value as the tolerance
                tolerance = min(deviation_values)
            else:
                # Cannot find tolerance for this shape
                tolerance = ''
        tolerance_dict[f'Dimension_Type_{shape}_Tolerance'] = tolerance
        count_tolerance_dict = {**count_dict, **tolerance_dict}
    return count_tolerance_dict


# Function to count the total number of sectional views present
def count_sectional_directories(root_dir):
    total_count = 0

    # loop through all subdirectories in root_dir
    for page_dir in os.listdir(root_dir):
        page_path = os.path.join(root_dir, page_dir)
        if os.path.isdir(page_path):
            for sheet_dir in os.listdir(page_path):
                sheet_path = os.path.join(page_path, sheet_dir)
                if os.path.isdir(sheet_path):
                    for canvas_dir in os.listdir(sheet_path):
                        canvas_path = os.path.join(sheet_path, canvas_dir)
                        if os.path.isdir(canvas_path) and canvas_dir.startswith('Canvas'):
                            for sectional_dir in os.listdir(canvas_path):
                                sectional_path = os.path.join(canvas_path, sectional_dir)
                                if os.path.isdir(sectional_path) and sectional_dir.startswith('Sectional'):
                                    total_count += 1
    return total_count



# Function to Calculate the Inclined_Drilling_Count and Inclined_Drilling_Values
def get_drilling_count_values(df):
    Inclined_Drilling_Count = 0
    Inclined_Drilling_Values = str()

    CUTOFF_INCLINED_DRILLING_POSITION_DIFFERENCE = 4
    label_column = get_coulmn_data(df, 'Measures', 1)
    shape_column = get_coulmn_data(df, 'Measures', 5)
    sectional_position_column = get_coulmn_data(df, 'Measures', 3)

    # Converting list of sectional_position_column to list of numerical tuples for calculations
    sectional_position_result = []
    for sectional_position in sectional_position_column:
        # remove the parentheses and split the string by comma
        nums = sectional_position.replace('(', '').replace(')', '').split(',')
        # convert the strings to integers and create a tuple
        tup = (int(nums[0]), int(nums[1])), (int(nums[2]), int(nums[3]))
        # append the tuple to the result list
        sectional_position_result.append(tup)

    # Looping through shape_column to find 'ROUND' shape
    for index in range(len(shape_column)):
        if shape_column[index] == 'ROUND':
            sectional_position_index = sectional_position_result[index]
            x1 = sectional_position_index[0][0]
            y1 = sectional_position_index[0][1]
            x2 = sectional_position_index[1][0]
            y2 = sectional_position_index[1][1]

            if abs(x1 - x2) > CUTOFF_INCLINED_DRILLING_POSITION_DIFFERENCE and abs(
                    y1 - y2) > CUTOFF_INCLINED_DRILLING_POSITION_DIFFERENCE:
                Inclined_Drilling_Count += 1
                Inclined_Drilling_Values = Inclined_Drilling_Values + label_column[index] + ','


    # Removing the last ',' from the resultant Inclined_Drilling_Values
    Inclined_Drilling_Values = Inclined_Drilling_Values[:-1]
    return Inclined_Drilling_Count, Inclined_Drilling_Values





if __name__ == '__main__':

    parser = argparse.ArgumentParser()
    parser.add_argument('--root_folder', help='Path for the root directory. If not provided, system will read folder from working directory')
    parser.add_argument('--output_loc', help='Path to store output csv file. If not provided, system will save csv to  working directory')
    args = parser.parse_args()

    if(args.root_folder):
        root_folder = args.root_folder
    else:
        root_folder = os.path.join(os.getcwd(), 'Results_Task_edited')



    if(args.output_loc):
        output_loc = args.output_loc+'/Challenge_Data_Analyst_Harshit_Tyagi.csv'
    else:
        output_loc = os.getcwd()+'/Challenge_Data_Analyst_Harshit_Tyagi.csv'


    output_list = []

    headers = ['Result', 'Drawing_Index', 'Drawing_Type', 'Dp_Number', 'Drawing_Id', 'Weight',
               'Dimension_Type_Nominal_Count', 'Dimension_Type_Nominal_Tolerance',
               'Dimension_Type_Round_Count', 'Dimension_Type_Round_Tolerance', 'Inclined_Drilling_Count',
               'Inclined_Drilling_Values', 'View_Count']



        # Loop through all the sub-folders
    for subdir in os.listdir(root_folder):

        sub_path = os.path.join(root_folder, subdir)
        total_sectional_view = count_sectional_directories(sub_path)
        directory, Drawing_Index, Drawing_Type, Dp_Number = extract_drawing_info(subdir)

        if os.path.isdir(sub_path):
            dfs = []

            # Loop through all the Page sub-folders
            for page_dir in os.listdir(sub_path):
                page_path = os.path.join(sub_path, page_dir)

                if os.path.isdir(page_path):

                    # Get the path to the excel file
                    excel_path = os.path.join(page_path, "Human Friendly Results")
                    if os.path.isfile(excel_path + ".xls"):
                        excel_path += ".xls"
                    elif os.path.isfile(excel_path + ".xlsx"):
                        excel_path += ".xlsx"
                    else:
                        continue

                    # Read the excel file and append to the list of dataframes
                    df = pd.read_excel(excel_path, sheet_name=None)
                    dfs.append(df)

            # Calling the functions to get the values
            drawing_id = get_drawing_id(dfs)
            weight = get_weight(dfs)
            shape_count = get_shape_count(dfs)
            Inclined_Drilling_Count, Inclined_Drilling_Values = get_drilling_count_values(dfs)

            dict_result = {'Result':directory,
                               'Drawing_Index':Drawing_Index,
                               'Drawing_Type':Drawing_Type,
                               'Dp_Number':Dp_Number,
                               'Drawing_Id':drawing_id,
                               'Weight':weight,
                                'Dimension_Type_Nominal_Count':shape_count.get('Dimension_Type_NOMINAL_Count'),
                               'Dimension_Type_Nominal_Tolerance':shape_count.get('Dimension_Type_NOMINAL_Tolerance'),
                                'Dimension_Type_Round_Count':shape_count.get('Dimension_Type_ROUND_Count'),
                               'Dimension_Type_Round_Tolerance':shape_count.get('Dimension_Type_ROUND_Tolerance'),
                                'Inclined_Drilling_Count':Inclined_Drilling_Count,
                                'Inclined_Drilling_Values':Inclined_Drilling_Values,
                               'View_Count':total_sectional_view}

            output_list.append(dict_result)

    # Sort the resultant dictionary on the basis of 'Result' column
    sorted_output_list = sorted(output_list, key=lambda x: x['Result'])

    with open(output_loc, 'w', newline='') as csvfile:
        # Create a DictWriter object
        writer = csv.DictWriter(csvfile, fieldnames=headers)

        # Write the headers to the CSV file
        writer.writeheader()

        # Iterate over the sorted data and write each row to the CSV file
        for row in sorted_output_list:
            writer.writerow(row)

    csvfile.close()


