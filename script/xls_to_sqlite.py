import xlrd
import glob
import os
import sys
import sqlite3 as sqlite

# convert the spreadsheets to an SQLite database

# The XLS file contains these columns:
# EPA Facility ID,	Facility Name,	Facility Address,	City,	State,	County,	Zip,	Facility DUNS,	Latitude,	Longitude,	Chemical(s),	NAICS Code(s)
#
# Create these base tables:
#
# 1. facilities, containing EPA Facility ID,	Facility Name,	Facility Address,	City,	State,	County,	Zip,	Facility DUNS,	Latitude,	Longitude
# 2. chemicals, containing the list of chemicals
# 3. naics_codes, containing the list of NAICS codes
#
# create the following join tables:
#
# 1. facilities_chemicals: many-to-many table of facility id to chemical id
# 2. facilities_nacis_codes: many-to-many table of facility id to naics_code id

# create the databases


def create_facilities_table(cursor):
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS facilities (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            epa_facility_id TEXT,
            facility_name TEXT,
            facility_address TEXT,
            city TEXT,
            state TEXT,
            county TEXT,
            zip TEXT,
            facility_duns TEXT,
            latitude REAL,
            longitude REAL
        )
    """
    )


def create_chemicals_table(cursor):
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS chemicals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chemical TEXT
        )
    """
    )


def create_naics_codes_table(cursor):
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS naics_codes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            naics_code TEXT
        )
    """
    )


def create_facilities_chemicals_table(cursor):
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS facilities_chemicals (
            facility_id INTEGER,
            chemical_id INTEGER
        )
    """
    )


def create_facilities_naics_codes_table(cursor):
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS facilities_naics_codes (
            facility_id INTEGER,
            naics_code_id INTEGER
        )
    """
    )


def create_tables(cursor):
    create_facilities_table(cursor)
    create_chemicals_table(cursor)
    create_naics_codes_table(cursor)
    create_facilities_chemicals_table(cursor)
    create_facilities_naics_codes_table(cursor)


def drop_tables(cursor):
    cursor.execute(
        """
        DROP TABLE IF EXISTS facilities
    """
    )
    cursor.execute(
        """
        DROP TABLE IF EXISTS chemicals
    """
    )
    cursor.execute(
        """
        DROP TABLE IF EXISTS naics_codes
    """
    )
    cursor.execute(
        """
        DROP TABLE IF EXISTS facilities_chemicals
    """
    )
    cursor.execute(
        """
        DROP TABLE IF EXISTS facilities_naics_codes
    """
    )


def get_or_create_chemical_id(cursor, chemical):
    cursor.execute(
        """
        SELECT id FROM chemicals WHERE chemical = ?
    """,
        (chemical,),
    )

    row = cursor.fetchone()

    if row:
        return row[0]

    cursor.execute(
        """
        INSERT INTO chemicals (
            chemical
        )
        VALUES (?)
    """,
        (chemical,),
    )

    return cursor.lastrowid


def get_or_create_naics_code_id(cursor, naics_code):
    cursor.execute(
        """
        SELECT id FROM naics_codes WHERE naics_code = ?
    """,
        (naics_code,),
    )

    row = cursor.fetchone()

    if row:
        return row[0]

    cursor.execute(
        """
        INSERT INTO naics_codes (
            naics_code
        )
        VALUES (?)
    """,
        (naics_code,),
    )

    return cursor.lastrowid


# given a row from the spreadsheet, insert the data into the database
# and return the facility id
#
# The XLS file contains these columns:
# EPA Facility ID,	Facility Name,	Facility Address,	City,	State,	County,	Zip,	Facility DUNS,	Latitude,	Longitude,	Chemical(s),	NAICS Code(s)
# The chemical and naics codes are comma separated lists
#
# The facility id is the first column
# The facility name is the second column
# The facility address is the third column
# The city is the fourth column
# The state is the fifth column
# The county is the sixth column
# The zip is the seventh column
# The facility duns is the eighth column
# The latitude is the ninth column
# The longitude is the tenth column
# The chemicals are in the eleventh column
# The naics codes are in the twelfth column
def insert_facility(cursor, row):
    epa_facility_id = row[0]
    facility_name = row[1]
    facility_address = row[2]
    city = row[3]
    state = row[4]
    county = row[5]
    zip = row[6]
    facility_duns = row[7]
    latitude = row[8]
    longitude = row[9]
    chemicals = row[10].split(",")
    naics_codes = row[11].split(",")

    cursor.execute(
        """
        INSERT INTO facilities (
            epa_facility_id,
            facility_name,
            facility_address,
            city,
            state,
            county,
            zip,
            facility_duns,
            latitude,
            longitude
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """,
        (
            epa_facility_id,
            facility_name,
            facility_address,
            city,
            state,
            county,
            zip,
            facility_duns,
            latitude,
            longitude,
        ),
    )

    facility_id = cursor.lastrowid

    for chemical in chemicals:
        if not chemical.strip():
            continue
        chemical_id = get_or_create_chemical_id(cursor, chemical.strip())
        cursor.execute(
            """
            INSERT INTO facilities_chemicals (
                facility_id,
                chemical_id
            )
            VALUES (?, ?)
        """,
            (facility_id, chemical_id),
        )

    for naics_code in naics_codes:
        # insert only if the naics code is not empty and isn't already in the database
        if not naics_code.strip():
            continue
        naics_code_id = get_or_create_naics_code_id(cursor, naics_code.strip())
        cursor.execute(
            """
            INSERT INTO facilities_naics_codes (
                facility_id,
                naics_code_id
            )
            VALUES (?, ?)
        """,
            (facility_id, naics_code_id),
        )

    return facility_id


def process_xls_files(cursor, xls_files):
    for xls_file in xls_files:
        print(f"Processing {xls_file}", file=sys.stderr, flush=True, end=" ... ")
        workbook = xlrd.open_workbook(xls_file)
        sheet = workbook.sheet_by_index(0)

        for i in range(1, sheet.nrows):
            row = sheet.row_values(i)
            insert_facility(cursor, row)
        print("done", file=sys.stderr, flush=True)


def main():
    # create the database
    conn = sqlite.connect("facilities.db")
    cursor = conn.cursor()

    # drop the tables
    drop_tables(cursor)

    # create the tables
    create_tables(cursor)

    # get the XLS files
    xls_files = glob.glob("reports/*.xls")

    # process the XLS files
    process_xls_files(cursor, xls_files)

    # commit the changes
    conn.commit()

    # close the connection
    conn.close()


if __name__ == "__main__":
    main()
