import xlrd
import xlwt
import glob
import os
import sys


def combine_xls_files(input_folder, output_file):
    # Create a new workbook and add a worksheet
    output_workbook = xlwt.Workbook()
    output_sheet = output_workbook.add_sheet("Combined")

    # Initialize row index for the output sheet
    output_row_index = 0

    # Get all .xls files in the input folder
    xls_files = glob.glob(os.path.join(input_folder, "*.xls"))

    for file_index, xls_file in enumerate(xls_files):
        # Open the workbook and select the first sheet
        workbook = xlrd.open_workbook(xls_file)
        sheet = workbook.sheet_by_index(0)

        # Iterate through the rows in the sheet
        for row_index in range(sheet.nrows):
            # Skip the header row for all files except the first one
            if file_index > 0 and row_index == 0:
                continue

            # Write the row to the output sheet
            for col_index in range(sheet.ncols):
                output_sheet.write(
                    output_row_index, col_index, sheet.cell_value(row_index, col_index)
                )

            # Increment the output row index
            output_row_index += 1

    # Save the combined workbook
    output_workbook.save(output_file)


# Usage
# set input_folder to argument 1 and output_file to argument 2
input_folder = sys.argv[1]
output_file = sys.argv[2]
combine_xls_files(input_folder, output_file)
