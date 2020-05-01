import csv
import json
import operator
import requests
import sys
import typing

# ---------------------------
#         CONSTANTS
# ---------------------------
BUILD = "Build"
EXTERNAL = "External"
HOT_FIX = "Hotfix"
IDEMPOTENT = "Idempotent"
INTERNAL = 'Internal'
MAJOR = "MajorVersion"
METHOD_ACCESS = "MethodAccess"
METHOD_ACCESS_CHECKED = "MethodAccessChecked"
METHOD_CLASS = "MethodClass"
METHOD_LIST = 'MethodList'
MINOR = "MinorVersion"
NAME = 'Name'
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
        print(f"Column Mismatch: {diff}")
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
        sort_key: (optional) sort the rows based on one of the header elements (DEFAULT: Name)

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


# +------------------------------------+
# |           APPLICATION              |
# +------------------------------------+
if __name__ == "__main__":

    # API URL for list of all APIs
    api_url = 'https://price.pclender.com/nexbank/method_list'

    # Expected columns
    expected_cols = [NAME, IDEMPOTENT, RESULT_TYPE, METHOD_ACCESS_CHECKED, METHOD_CLASS, INTERNAL, METHOD_ACCESS]

    # Get the data (dictionary) from the PRICE URL
    full_data = get_api_list(url=api_url)
    raw_api_data = full_data.get(METHOD_LIST)

    # Get the list of columns listed in the data
    columns = [col for col in raw_api_data[0].keys()]

    # Build lists of API dictionaries based on INTERNAL or EXTERNAL data (EXTERNAL = not INTERNAL)
    api_data = {
        EXTERNAL:  [rows for rows in raw_api_data if not rows.get(INTERNAL, False)],
        INTERNAL: [rows for rows in raw_api_data if rows.get(INTERNAL, False)],
    }

    # Verify data columns contain the expected columns
    if not verify_cols_are_present(source=columns, expected=expected_cols):
        print("Unable to generate files... exiting.")
        sys.exit()

    # Generate the CSV files for INTERNAL and EXTERNAL APIs.
    print(f"All expected columns present, generating CSVs.")
    for data_type, api_dict in api_data.items():
        write_csv_from_dict(filename=f"{data_type.lower()}_apis.csv",
                            ordered_column_header=[c for c in expected_cols if c != INTERNAL],
                            data_list=api_dict)

    # Print out the version for reference.
    print(f"VERSION: {build_version_str(full_data.get(VERSION))}")
