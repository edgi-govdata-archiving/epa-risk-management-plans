import xlrd
import glob
import os
import time

# Generate a Markdown file with links to the reports.
# Read all the XLS files in the reports dirrectory
# The XLS file contains these columns:
# EPA Facility ID,	Facility Name,	Facility Address,	City,	State,	County,	Zip,	Facility DUNS,	Latitude,	Longitude,	Chemical(s),	NAICS Code(s)
# For each XLS file, get the state name from the filename, and use that as a second-level heading.
# Then, create a Markdown table with the following columns:
# EPA Facility ID, Facility Name, Facility Address, City, County, Zip, Facility DUNS, Latitude, Longitude, Chemical(s), NAICS Code(s)
# Create a link to the PDF report in the EPA Facility ID column, if the PDF file exists.


def state_name_to_directory_name(state: str) -> str:
    return state.replace(" ", "_").lower()


def create_state_index_md(state: str, xls_file: str, data_dict: dict):
    state_directory = state_name_to_directory_name(state)
    full_state_directory = f"reports/{state_directory}"
    if not os.path.exists(full_state_directory):
        return ""
    xls_creation_time = os.path.getctime(xls_file)
    xls_creation_date = time.ctime(xls_creation_time)

    book = xlrd.open_workbook(xls_file)
    sheet = book.sheet_by_index(0)
    with open(f"{full_state_directory}/index.md", "w") as f:
        f.write(f"# {state}\n\n")
        f.write(f"The state data was downloaded on ({xls_creation_date})\n\n")
        n, completed = data_dict[state_directory]
        percent_completed = float(completed) / n if n > 0 else 0
        f.write(
            f"Of the {n} reports, {completed} were downloaded ({percent_completed:.2%})\n\n"
        )
        f.write(
            "| EPA Facility ID | Facility Name | Facility Address | City | County | Zip | Facility DUNS | Latitude | Longitude | Chemical(s) | NAICS Code(s) | Update Date |\n"
        )
        f.write(
            "|------------------|---------------|------------------|------|--------|-----|--------------|----------|-----------|--------------|---------------|------------|\n"
        )
        for row in range(1, sheet.nrows):
            row_data = sheet.row_values(row)
            for row in range(1, sheet.nrows):
                row_data = sheet.row_values(row)
                facility_id = row_data[0]
                facility_name = row_data[1]
                facility_address = row_data[2]
                city = row_data[3]
                county = row_data[5]
                zip = row_data[6]
                facility_duns = row_data[7]
                latitude = row_data[8]
                longitude = row_data[9]
                chemicals = row_data[10]
                naics_codes = row_data[11]
                pdf_file = f"reports/{facility_id}.pdf"
                fac_creation_date = "N/A"
                fac_text = facility_id
                if os.path.exists(pdf_file):
                    fac_text = f" [{facility_id}](./{pdf_file})"
                    fac_creation_time = os.path.getctime(pdf_file)
                    fac_creation_date = time.ctime(fac_creation_time)
                f.write(
                    f"| {fac_text} | {facility_name} | {facility_address} | {city} | {county} | {zip} | {facility_duns} | {latitude} | {longitude} | {chemicals} | {naics_codes} | {fac_creation_date}\n"
                )


def count_completed_reports(
    state: str, verbose: bool = False
) -> (str, int, int, float):
    state_directory = state_name_to_directory_name(state)
    full_state_directory = f"reports/{state_directory}"
    if not os.path.exists(full_state_directory):
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
        pdf_file = f"{full_state_directory}/{facility_id}.pdf"
        if os.path.exists(pdf_file):
            completed += 1
        else:
            if verbose:
                print(f"Missing report for {state} facility {facility_id}")
    if n == 0:
        return (state_directory, 0, 0, 0.0)
    return (state_directory, completed, n)


def report_stats():
    # get a list of 'xls' files from reports directory
    xls_files = glob.glob("reports/*.xls")
    xls_files.sort()
    total = 0
    total_completed = 0
    data_dict = {}
    for xls_file in xls_files:
        state = os.path.basename(xls_file).split(".")[0]
        directory = state_name_to_directory_name(state)
        _, completed, n = count_completed_reports(state)
        total += n
        total_completed += completed
        data_dict[directory] = (completed, n)
    return data_dict, total, total_completed


if __name__ == "__main__":
    # get a list of 'xlsx' files from reports directory
    xls_files = glob.glob("reports/*.xls")
    # sort by state name
    xls_files.sort()
    data_dict, total, total_completed = report_stats()
    with open("index.md", "w") as f:
        f.write("# EPA RMP Data\n\n")
        f.write(
            f"Of the {total} reports, {total_completed} were downloaded ({total_completed/total:.2%})\n\n"
        )
        for xls_file in xls_files:
            state = os.path.basename(xls_file).split(".")[0]
            print(f"Creating index for {state} ... ", end="", flush=True)
            create_state_index_md(state, xls_file, data_dict)
            f.write(f"- [{state}](./reports/{state}/index.md)\n")
            print("done.")
