from collections import OrderedDict
import csv
import json
import operator
import requests
import sys
import typing

import xlsxwriter


# ---------------------------
#         CONSTANTS
# ---------------------------
BUILD = "Build"
EXTERNAL = "External"
FISERV = "Fiserv, Inc."
HOT_FIX = "Hotfix"
IDEMPOTENT = "Idempotent"
INTERNAL = "Internal"
MAJOR = "MajorVersion"
METHOD_ACCESS = "MethodAccess"
METHOD_ACCESS_CHECKED = "MethodAccessChecked"
METHOD_CLASS = "MethodClass"
METHOD_LIST = "MethodList"
MINOR = "MinorVersion"
NAME = "Name"
NOT_SET = "NOT FOUND"
RESULT_TYPE = "ResultType"
VERSION = "Version"


def build_version_str(version_info: typing.Dict[str, str]) -> str:
    """
    Builds the version string from the API response data (dict)
    Args:
        version_info: dictionary of major, minor, hotfix, and build elements.

    Returns:
        str: <MAJOR>.<MINOR>.<HOT_FIX>.<BUILD>
    """
    version_elements = [MAJOR, MINOR, HOT_FIX, BUILD]
    return ".".join([str(version_info.get(elem)) for elem in version_elements])


def get_api_list(url: str) -> typing.Dict[str, typing.Any]:
    """
    Call the specified API and return the response payload as JSON.

    Returns: (dict) JSON response, if the response code was 2xx.

    """
    response = requests.get(url)
    if int(response.status_code) / 100 != 2:
        print(f"Unable to get API information. Rec'd status code: {response.status_code}")
        return {}

    return json.loads(response.text)


def verify_cols_are_present(source: typing.List[str], expected: typing.List[str]) -> bool:
    """
    Verify the expected data columns were returned in the response.
    Args:
        source: List of columns returned in the response data
        expected: List of expected/required columns

    Returns:
        Boolean: True = Columns present, False = Actual columns list does not match the expected list (case sensitive).

    """
    diff = set(expected) - set(source)
    if diff:
        print(f"ERROR:\n\tColumn Mismatch(es): {diff}")
    return not diff


def write_csv_from_dict(
        filename: str, ordered_column_header: typing.List[str], data_list: typing.List[dict],
        sort_key=NAME) -> typing.NoReturn:
    """
    Writes the dictionary for each API in the list to a CSV file, using the expected column headers in the
    provided order.

    Args:
        filename: Name of file to write
        ordered_column_header: list of columns to write to the CSV file.
        data_list: list where each element defines a dictionary describing an API
        sort_key: (optional) sort the rows based on one of the header elements (DEFAULT: NAME Constant)

    Returns:
        None.
    """

    # Open the file, write the header row, and then the data.
    with open(filename, "w", newline='') as FILE:
        writer = csv.writer(FILE)
        writer.writerow(ordered_column_header)

        # Sort the data by the sort_key and write to file.
        for row in sorted(data_list, key=operator.itemgetter(sort_key)):
            writer.writerow([row.get(col) for col in ordered_column_header])

    print(f"Wrote file: {filename}")


class ExcelFile:

    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"

    def __init__(self, filename: str) -> typing.NoReturn:
        """
        Instantiates the XSLX file object.
            NOTE: obj.close() needs to be called for the file to be written.

        Args:
            filename: Name of XLSX to create

        """
        # filename: name of excel spreadsheet to create/use.
        self.filename = filename
        self._workbook = xlsxwriter.Workbook(filename)

    def set_workbook_properties(
            self, title: str = "", subject: str = "", author: str = "", comments: str = "", keywords: str = "",
            category: str = "", company: str = FISERV, status: str = "", manager: str = "",
            hyperlink_base: str = "") -> typing.NoReturn:
        """
        Set the workbook properties/metadata.

        Args:
            title: Name of workbook.
            subject: Topic or subject of workbook
            author: Author of workbook
            comments: Comments
            keywords: Keywords (for searching)
            category: Category of spreadsheet
            company: Company that created/owns spreadsheet
            status: Status of spreadsheet (e.g. initial draft, in progress, in review, complete, finalized)
            manager: Owner or responsible party for maintaining spreadsheet
            hyperlink_base:

        Returns:
            None
        """

        properties = {
            'title': title, 'subject': subject, 'author': author, 'comments': comments, 'keywords': keywords,
            'category': category, 'company': company, 'status': status, 'manager': manager,
            'hyperlink_base': hyperlink_base
        }

        self._workbook.set_properties(properties)

    def _get_unique_worksheet_name(self, worksheet_name: str) -> str:
        """
        Iterate through existing worksheet names, and if provided worksheet_name already exists,
        add proper index to worksheet_name.

        e.g. Provided worksheet name "my_wks" already exists, so the next wks name would be "my_wks_1"

        Args:
            worksheet_name: Name of worksheet

        Returns:
            If already unique: provided name of worksheet
            If worksheet name exists in the workbook: worksheet_{index} <<--- will be unique
        """

        # Get list of existing worksheets
        existing_wks = [wks.get_name() for wks in self._workbook.worksheets()]
        unique_wks_name = worksheet_name
        index = 1

        # Until a unique worksheet name is found...
        while unique_wks_name in existing_wks:
            unique_wks_name = f"{worksheet_name}_{index}"
            index += 1
        return unique_wks_name

    def close_workbook(self) -> typing.NoReturn:
        """
        Closes the workbook and writes the workbook to file.
        Returns:
            None
        """
        self._workbook.close()
        print(f"Created XLSX File: {self.filename}")

    def create_worksheet(self, column_alignment_dict: typing.Dict[str, str], data_list: typing.List[dict],
                         sort_key: str = NAME, worksheet_name: str = '') -> typing.NoReturn:
        """
        Write data to the excel file. If the file exists, it will add a tab to the file.
        The initial file creation will create a tab that matches the filename (if the tab name is not provided).

        Args:

            column_alignment_dict: an ordered list of columns (to be the sheet header (row 0),
                    and used to print out each row in the correct order.
            sort_key: Column to sort data by (written to worksheet in this order).
            data_list: the rows of data to write to the spreadsheet tab (will be specific to this app)
            worksheet_name: Name of the excel worksheet for the data

        Returns:
            None

        """
        column_width = {}
        column_buffer = 7  # Add a little bit of margin due to font kerning (not all letters take up the same space)

        worksheet_name = worksheet_name if worksheet_name != '' else self.filename

        # Create the worksheet
        unique_worksheet_name = self._get_unique_worksheet_name(worksheet_name=worksheet_name)
        worksheet = self._workbook.add_worksheet(unique_worksheet_name)

        # Populate and format the header row
        for column_index, column_name in enumerate(column_alignment_dict.keys()):

            # Align the column accordingly and make the entries bold.
            column_format = self._workbook.add_format(
                {'align': column_alignment_dict[column_name], 'bold': True})

            # Set header background to light gray
            column_format.set_bg_color("D9D9D9")
            worksheet.write(0, column_index, column_name, column_format)
            column_width[column_name] = len(column_name)

        # Freeze the top pane (header row) of the spreadsheet
        worksheet.freeze_panes(1, 0)

        # Populate the remainder of the worksheet, sorted by sort_key, starting at row 2 (row 1= header).
        # Also track the length of all entries for each column, so that the columns can be resized correctly.
        for row_index, row_data in enumerate(sorted(data_list, key=operator.itemgetter(sort_key)), start=1):
            for column_index, column_name in enumerate(column_alignment_dict):
                entry = str(row_data.get(column_name))
                worksheet.write_string(row_index, column_index, entry)
                if len(entry) > column_width[column_name]:
                    column_width[column_name] = len(entry)

        # Adjust each column width to support longest entry
        for column_index, column_name in enumerate(column_alignment_dict):
            column_format = self._workbook.add_format({'align': column_alignment_dict[column_name]})
            worksheet.set_column(column_index, column_index, column_width[column_name] + column_buffer, column_format)

        print(f"Created and populated an excel worksheet: {unique_worksheet_name}")


# +------------------------------------+
# |           APPLICATION              |
# +------------------------------------+
if __name__ == "__main__":

    # API URL for list of all APIs
    api_url = 'https://price.pclender.com/nexbank/method_list'

    # File formats to create
    CREATE_CSV = False
    CREATE_XLSX = True

    # Expected Data Columns (in desired displayed order) and their alignments
    column_alignment = OrderedDict()
    column_alignment[NAME] = ExcelFile.LEFT
    column_alignment[IDEMPOTENT] = ExcelFile.CENTER
    column_alignment[RESULT_TYPE] = ExcelFile.LEFT
    column_alignment[METHOD_ACCESS_CHECKED] = ExcelFile.CENTER
    column_alignment[METHOD_CLASS] = ExcelFile.LEFT
    column_alignment[INTERNAL] = ExcelFile.CENTER
    column_alignment[METHOD_ACCESS] = ExcelFile.LEFT

    # Get the data (dictionaries) from the PRICE URL
    full_data = get_api_list(url=api_url)
    raw_api_data = full_data.get(METHOD_LIST)
    version = build_version_str(full_data.get(VERSION))

    # Build lists of API dictionaries based on INTERNAL or EXTERNAL data (NOTE: EXTERNAL = not INTERNAL)
    api_data = {
        EXTERNAL:  [rows for rows in raw_api_data if not rows.get(INTERNAL, False)],
        INTERNAL: [rows for rows in raw_api_data if rows.get(INTERNAL, False)],
    }

    # Verify data columns contain the expected columns
    if not verify_cols_are_present(
            source=[col for col in raw_api_data[0].keys()],
            expected=list(column_alignment.keys())):
        print("Unable to generate files, data columns do not match expected columns...\nExiting.\n")
        sys.exit()

    print("All expected columns present, generating designated files.")

    # If XLSX, create the workbook and set the properties.
    if CREATE_XLSX:
        xlsx_workbook = ExcelFile(filename=f"API_Mapping_{version}.xlsx")
        xlsx_workbook.set_workbook_properties(
            title=f"PRICE API Mapping: {version}", subject="PRICE APIs", keywords="API",
            comments="Generated by Fiserv Product Automation", status="In Production", author="Python Automation")

    # No longer needed since the data is segregated by this value.
    del column_alignment[INTERNAL]

    # Generate the output files for INTERNAL and EXTERNAL APIs.
    for data_type, api_dict in api_data.items():

        if CREATE_CSV:
            write_csv_from_dict(
                filename=f"{data_type.lower()}_apis.csv", ordered_column_header=list(column_alignment.keys()),
                data_list=api_dict)

        if CREATE_XLSX:
            xlsx_workbook.create_worksheet(
                column_alignment_dict=column_alignment, data_list=api_dict, worksheet_name=data_type)

    # If building an XLSX file: add the version worksheet where the name = version number
    if CREATE_XLSX:
        xlsx_workbook.create_worksheet(column_alignment_dict={"Version": ExcelFile.LEFT},
                                       data_list=[{VERSION: version}], worksheet_name=version, sort_key=VERSION)
        xlsx_workbook.close_workbook()

    # Print out the version for reference.
    print(f"Files generated for API VERSION: {version}")