import os
import sys

from pandas import read_excel, ExcelWriter
import argparse


def set_names(in_name, out_name, checksame=False):
    """Ensure input and output names have correct extensions.
    Optionally check if they are the same and abort if they are.

    Args:
        in_name (String): Input filename
        out_name (String): Output filename
        checksame (bool, optional): Toggle whether to check if input and 
        output are the same. Defaults to False.

    Returns:
        (String, String): The resultant input and output names
    """
    if not in_name.endswith(".xlsx"):
        in_name += ".xlsx"

    if not out_name.endswith(".xlsx"):
        out_name += ".xlsx"

    if (out_name == in_name) and checksame:
        sys.exit("Input name matches output name, aborting")

    return in_name, out_name


def clean_sheet(input_name):
    """Run the cleaner over the spreadsheet

    Args:
        input_name (String): Filename of the input spreadsheet

    Returns:
        pandas.DataFrame: Dataframe of the cleaned spreadsheet
    """
    # Load the spreadsheet
    df = read_excel(input_name, header=None, index_col=None)

    # Replace any whitespace with empty strings
    df = df.replace(r"^\s*$", "", regex=True)
    return df


def output_sheet(df, output_name):
    """Output the dataframe as a spreadsheet.

    Args:
        df (pandas.DataFrame): Dataframe to be output
        output_name (String): Output filename
    """
    writer = ExcelWriter(output_name, engine="xlsxwriter")
    df.to_excel(writer, sheet_name="Sheet1", index=False, header=False)
    worksheet = writer.sheets["Sheet1"]
    writer.save()


if __name__ == "__main__":
    # Load arguments
    parser = argparse.ArgumentParser(
        description="Clean an excel sheet such that all" +
        " cells containing whitespace or empty strings are blank")

    parser.add_argument("input", help="Name of input file")
    parser.add_argument("output", help="Name of output file")

    args = parser.parse_args()
    input_name = args.input
    output_name = args.output

    input_name, output_name = set_names(input_name,
                                        output_name,
                                        checksame=True)

    print(f"Loading file: {input_name}")
    df = clean_sheet(input_name)

    # Output new sheet
    print(f"Writing {output_name}")
    output_sheet(df, output_name)
    print("Done")
