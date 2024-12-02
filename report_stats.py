import xlrd
import glob
import os
import sys
import sqlite3 as sqlite
from os import linesep


def dedent(message):
    return linesep.join(line.lstrip() for line in message.splitlines())


def count_total(cursor):
    cursor.execute(
        """
        SELECT COUNT(*)
        FROM facilities
        """
    )
    return cursor.fetchone()[0]


def count_distinct_facilities(cursor):
    cursor.execute(
        """
        SELECT COUNT(DISTINCT facility_name)
        FROM facilities
        """
    )
    return cursor.fetchone()[0]


def count_by_state(cursor):
    cursor.execute(
        """
        SELECT state, COUNT(*)
        FROM facilities
        GROUP BY state
        ORDER BY count(*) DESC, state
        """
    )
    return cursor.fetchall()


def count_by_county_and_state(cursor, n=100):
    cursor.execute(
        """
        SELECT state, county, COUNT(*)
        FROM facilities
        GROUP BY state, county
        ORDER BY count(*) DESC, county, state
        LIMIT ?
        """,
        (n,),
    )
    return cursor.fetchall()


def count_by_naics_code(cursor, n=100):
    cursor.execute(
        """
        SELECT naics_code, COUNT(*)
        FROM facilities_naics_codes
        JOIN naics_codes
        ON naics_code_id = naics_codes.id
        GROUP BY naics_code
        ORDER BY COUNT(*) DESC, naics_code
        LIMIT ?
        """,
        (n,),
    )
    return cursor.fetchall()


def count_by_chemical(cursor, n=100):
    cursor.execute(
        """
        SELECT chemical, COUNT(*)
        FROM facilities_chemicals
        JOIN chemicals
        ON chemical_id = chemicals.id
        GROUP BY chemical
        ORDER BY COUNT(*) DESC, chemical
        LIMIT ?
        """,
        (n,),
    )
    return cursor.fetchall()


def create_markdown_report(n=100):
    with sqlite.connect("facilities.db") as connection:
        cursor = connection.cursor()
        with open("stats.md", "w") as report:
            report.write("# Facilities Report\n\n")
            report.write(
                dedent(
                    f"""
                ## Basic Stats\n\n
                - Total Number of RMPs: {count_total(cursor):,}\n
                - Total Number of Facilities: {count_distinct_facilities(cursor):,}.\n\n
                """
                )
            )
            report.write("## Count by State\n\n")
            report.write("| State | Count |\n")
            report.write("|------:|------:|\n")
            for state, count in count_by_state(cursor):
                report.write(f"| {state} | {count:,} |\n")
            report.write("\n")
            report.write(f"## Count by County and State (top {n})\n\n")
            report.write("| County| Count |\n")
            report.write("|-------:|------:|\n")
            for state, county, count in count_by_county_and_state(cursor, n):
                report.write(f"| {county}, {state} | {count:,} |\n")
            report.write("\n")
            report.write(f"## Count by NAICS Code (top {n})\n\n")
            report.write("| NAICS Code | Count |\n")
            report.write("|-----------:|------:|\n")
            for naics_code, count in count_by_naics_code(cursor, n):
                report.write(f"| {naics_code} | {count:,} |\n")
            report.write("\n")
            report.write(f"## Count by Chemical (top {n})\n\n")
            report.write("| Chemical | Count |\n")
            report.write("|---------:|------:|\n")
            for chemical, count in count_by_chemical(cursor, n):
                report.write(f"| {chemical} | {count:,} |\n")
            report.write("\n")


if __name__ == "__main__":
    create_markdown_report()
