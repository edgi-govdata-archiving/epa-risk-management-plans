from playwright.sync_api import Page, sync_playwright
import time
import base64
import os
import random
import xlrd
import glob
import sys


def state_name_to_directory_name(state: str) -> str:
    return state.replace(" ", "_").lower()


def js_script(blob_url: str) -> str:
    return f"""
        const fetchBlob = async (url) => {{
            const response = await fetch(url);
            const blob = await response.blob();
            const reader = new FileReader();
            return new Promise((resolve, reject) => {{
                reader.onloadend = () => resolve(reader.result);
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            }});
        }};
        fetchBlob('{blob_url}');
        """


def scrape_report(
    page: Page,
    facility_id: str = "100000002043",
    state: str = "Michigan",
):

    page.goto("https://cdxapps.epa.gov/olem-rmp-pds/")
    page.get_by_label("EPA Facility ID.").fill(facility_id)
    page.locator("#state").get_by_role("combobox").click()
    page.get_by_role("option", name=state, exact=True).click()
    page.locator("pds-search-form").get_by_role("button", name="Search").click()

    with page.expect_popup() as popup:
        page.get_by_role("row", name=facility_id).get_by_role("button").click()
        popped = popup.value

        data_url = page.evaluate(js_script(blob_url=popped.url))
        header, data = data_url.split(",", 1)
        pdf_data = base64.b64decode(data)
        directory = state_name_to_directory_name(state)
        fname = f"reports/{directory}/{facility_id}.pdf"
        with open(fname, "wb") as f:
            f.write(pdf_data)


if __name__ == "__main__":
    # if the first argument is a state name, then only get the reports for that state
    # if the first argument looks like "-Texas", then get the reports for all states except Texas
    state = sys.argv[1] if len(sys.argv) > 1 else None
    if state and not state.startswith("-"):
        xls_files = glob.glob(f"reports/{state}.xls")
    else:
        # get a list of 'xlsx' files from reports directory
        xls_files = glob.glob("reports/*.xls")
        if state and state.startswith("-"):
            state = state[1:]
            xls_files = [xls_file for xls_file in xls_files if state not in xls_file]
    for state in xls_files:
        print(state)

    # randomize so we can start from any point, roughly
    random.shuffle(xls_files)

    for xls_file in xls_files:
        state = os.path.basename(xls_file).split(".")[0]
        directory = state_name_to_directory_name(state)
        print("Getting facility ids for state:", state, "... ", end="", flush=True)
        book = xlrd.open_workbook(xls_file)
        # facility ids are in the first column of the first sheet
        sheet = book.sheet_by_index(0)
        facility_ids = sheet.col_values(0)[1:]
        print("done.")
        # facility_ids = facility_ids[:1]
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            for facility_id in facility_ids:
                report_file = f"reports/{directory}/{facility_id}.pdf"
                if os.path.exists(report_file):
                    # print(f"Report already exists for {facility_id}")
                    continue
                print(
                    f"Getting data for {state} facility_id: {facility_id} ... ",
                    end="",
                    flush=True,
                )
                page = browser.new_page()
                try:
                    scrape_report(page, facility_id=facility_id, state=state)
                except Exception as e:
                    print(
                        f"UNABLE TO GET DATA FOR {state} FACILITY_ID: {facility_id} ERROR: {e}"
                    )
                    continue
                print("done.")
                time.sleep(random.randint(1, 5))
                page.close()
            browser.close()
