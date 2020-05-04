from collections import OrderedDict
import json
import operator
import requests
import sys
import typing

import xlsxwriter
from xlsxwriter.worksheet import Worksheet

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
URL = "URL"
VERSION = "Version"


# -------------------------------------------------------
#                  CLASSES
# -------------------------------------------------------
class ExcelFile:
    """ Used to build and populate an Excel file. (Currently implementation-specific to this application) """

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

    def _build_header(self, worksheet: Worksheet, column_dict: typing.Dict[str, str],
                      column_width_dict: typing.Dict[str, int]) -> typing.Dict[str, int]:
        """
        Builds and formats header row for specified worksheet based on entries in column_dict
        Args:
            worksheet: Worksheet to add header
            column_dict: Ordered Dictionary ==> k:column_names, v:default_alignments
            column_width_dict: dict for tracking longest entry in each column (used for resizing/auto-fit cols).

        Returns:
            column_width_dictionary (for use within worksheet population)
        """
        for column_index, column_name in enumerate(column_dict.keys()):

            # Align the column accordingly and make the entries bold.
            column_format = self._workbook.add_format(
                {'align': column_dict[column_name], 'bold': True})

            # Set header background to light gray
            column_format.set_bg_color("D9D9D9")

            worksheet.write(0, column_index, column_name, column_format)
            column_width_dict[column_name] = len(column_name)

        # Freeze the top pane (header row) of the spreadsheet
        worksheet.freeze_panes(1, 0)
        return column_width_dict

    def close_workbook(self) -> typing.NoReturn:
        """
        Closes the workbook and writes the workbook to file.
        Returns:
            None
        """
        self._workbook.close()
        print(f"Workbook: Wrote XLSX: {self.filename}")

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
        # Add a little bit of column margin due to font kerning (not all letters take up the same space).
        # Not needed for non-kerning fonts, such as Courier, but the font needs to be explicitly defined.
        # Currently, we are not defining the font, so Execl will default to Arial.
        column_buffer = 7

        worksheet_name = worksheet_name if worksheet_name != '' else self.filename

        # Create the worksheet
        unique_worksheet_name = self._get_unique_worksheet_name(worksheet_name=worksheet_name)
        worksheet = self._workbook.add_worksheet(unique_worksheet_name)

        # Builder worksheet header row
        column_width = self._build_header(
            worksheet=worksheet, column_dict=column_alignment_dict, column_width_dict={})

        # Populate the remainder of the worksheet, sorted by sort_key, starting at row 2 (row 1= header).
        # Also track the length of each entries for each column, so that each column can be resized correctly.
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

        print(f"Worksheet: Created/populated: {unique_worksheet_name}")


# -------------------------------------------------------
#             UTILITY-SPECIFIC ROUTINES
# -------------------------------------------------------
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
    payload = {}
    response = requests.get(url)
    if int(response.status_code) / 100 != 2:
        print(f"Unable to get API information. Rec'd status code: {response.status_code}")
    else:
        payload = json.loads(response.text)

    return payload


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


def define_expected_columns() -> OrderedDict:
    """
    Define the expected data columns (as provided by the API data); they are defined in the order
    that they should be presented in the output (e.g. - spreadsheet), and the expected alignment in the
    output file (e.g. - names are left justified, booleans are centered, etc.).

    Returns:
        Order Dictionary, where K: Column Name, V: alignment in output.

    """
    columns = OrderedDict()

    columns[NAME] = ExcelFile.LEFT
    columns[IDEMPOTENT] = ExcelFile.CENTER
    columns[RESULT_TYPE] = ExcelFile.LEFT
    columns[METHOD_ACCESS_CHECKED] = ExcelFile.CENTER
    columns[METHOD_CLASS] = ExcelFile.LEFT
    columns[INTERNAL] = ExcelFile.CENTER
    columns[METHOD_ACCESS] = ExcelFile.LEFT

    return columns


# -------------------------------------------------------
#           APPLICATION/EXECUTABLE
# -------------------------------------------------------
if __name__ == "__main__":

    # API URL for list of all APIs
    api_url = 'https://price.pclender.com/nexbank/method_list'
    if len(sys.argv) > 1:
        api_url = sys.argv[1]

    # Expected Data Columns (in desired displayed order) and their alignments
    column_alignment = define_expected_columns()

    # Get the data (dictionaries) from the PRICE URL
    full_data = get_api_list(url=api_url)
    raw_api_data = full_data.get(METHOD_LIST)
    version = build_version_str(full_data.get(VERSION))

    # Build lists of API dictionaries based on INTERNAL or EXTERNAL data (NOTE: EXTERNAL = not INTERNAL)
    api_data = {
        EXTERNAL: [rows for rows in raw_api_data if not rows.get(INTERNAL, False)],
        INTERNAL: [rows for rows in raw_api_data if rows.get(INTERNAL, False)],
    }

    # Verify data columns contain the expected columns
    if not verify_cols_are_present(
            source=[col for col in raw_api_data[0].keys()],
            expected=list(column_alignment.keys())):
        print("Unable to generate files, data columns do not match expected columns...\nExiting.\n")
        sys.exit()
    print("Validation: All expected columns present, generating requested files.")

    # No longer needed since the data is segregated by this value.
    del column_alignment[INTERNAL]

    # Create the workbook and set the properties.
    xlsx_workbook = ExcelFile(filename=f"API_Mapping_{version}.xlsx")
    xlsx_workbook.set_workbook_properties(
        title=f"PRICE API Mapping: {version}", subject="PRICE APIs", keywords="API",
        comments="Generated by Fiserv Product Automation", status="In Production", author="Python Automation")

    # Generate the output files for INTERNAL and EXTERNAL APIs.
    for data_type, api_dict in api_data.items():

        # Internal and external API worksheets
        xlsx_workbook.create_worksheet(
            column_alignment_dict=column_alignment, data_list=api_dict, worksheet_name=data_type)

    # Add the version worksheet where the name = version number
    xlsx_workbook.create_worksheet(column_alignment_dict={VERSION: ExcelFile.LEFT, URL: ExcelFile.LEFT},
                                   data_list=[{VERSION: version, URL: api_url}],
                                   worksheet_name=version, sort_key=VERSION)
    xlsx_workbook.close_workbook()

    # Print out the version for reference.
    print(f"COMPLETE: Generated using PRICE API VERSION: {version}")
