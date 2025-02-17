import xlrd
import glob
import os
import sys

# Count the number of reports in the reports directory, and see how many we have completed, by state
# The XLS file contains these columns:
# EPA Facility ID,	Facility Name,	Facility Address,	City,	State,	County,	Zip,	Facility DUNS,	Latitude,	Longitude,	Chemical(s),	NAICS Code(s)
# For each XLS file, get the state name from the filename, and use that as the directory name.
# Then count how many have a file in the reports directory.


def state_name_to_directory_name(state: str) -> str:
    return state.replace(" ", "_").lower()


def count_completed_reports(
    state: str, verbose: bool = False
) -> (str, int, int, float):
    state_directory = state_name_to_directory_name(state)
    state_directory = f"reports/{state_directory}"
    if not os.path.exists(state_directory):
        return 0
    xls_file = f"reports/{state}.xls"
    n = 0
    completed = 0
    book = xlrd.open_workbook(xls_file)
    sheet = book.sheet_by_index(0)
    for row in range(1, sheet.nrows):
        n += 1
        row_data = sheet.row_values(row)
        facility_id = row_data[0]
        pdf_file = f"{state_directory}/{facility_id}.pdf"
        if os.path.exists(pdf_file):
            completed += 1
        else:
            if verbose:
                print(f"Missing report for {state} facility {facility_id}")
    if n == 0:
        return (state, 0, 0, 0.0)
    return (state, completed, n, float(completed) / n)


if __name__ == "__main__":
    verbose = sys.argv[1] == "-v" if len(sys.argv) > 1 else False
    # get a list of 'xls' files from reports directory
    xls_files = glob.glob("reports/*.xls")
    # sort by state name
    xls_files.sort()
    total = 0
    total_completed = 0
    for xls_file in xls_files:
        state = os.path.basename(xls_file).split(".")[0]
        _, completed, n, percent = count_completed_reports(state, verbose)
        total += n
        total_completed += completed
        print(f"{state}: {completed}/{n} ({percent:.2%})")
    print(f"Total: {total_completed}/{total} ({total_completed/total:.2%})")
