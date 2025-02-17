from playwright.sync_api import Page, sync_playwright
import time
import base64
import os
import random


def scrape_excel(
    page: Page,
    state: str = "Michigan",
):
    page.goto("https://cdxapps.epa.gov/olem-rmp-pds/")
    page.locator("#state").get_by_role("combobox").click()
    page.get_by_role("option", name=state, exact=True).click()
    page.locator("pds-search-form").get_by_role("button", name="Search").click()
    with page.expect_download() as download_info:
        page.get_by_role("button", name="Download Results", exact=True).click()
        download = download_info.value
        fname = f"reports/{state}.xls"
        download.save_as(fname)


if __name__ == "__main__":
    # get a list of states from states.txt
    states = open("states.txt").read().splitlines()
    states = [f.strip() for f in states]
    # states = states[:1]
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        for state in states:
            report_file = f"reports/{state}.xls"
            if os.path.exists(report_file):
                print(f"Report already exists for {state}")
                continue
            print("Getting data for state:", state, "... ", end="", flush=True)
            page = browser.new_page()
            try:
                scrape_excel(page, state=state)
            except Exception as e:
                print("UNABLE TO GET DATA FOR STATE:", state, "ERROR:", e)
                continue
            print("done.")
            time.sleep(random.randint(1, 5))
            page.close()
        browser.close()
