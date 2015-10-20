#!/usr/bin/env python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import argparse

def letter_to_index(letter):
    """Converts a column letter, e.g. "A", "B", "AA", "BC" etc. to a zero based
    column index.

    http://stackoverflow.com/questions/13377793/is-it-possible-to-get-an-excel-documents-row-count-without-loading-the-entire-d

    A becomes 0, B becomes 1, Z becomes 25, AA becomes 26 etc.

    Args:
        letter (str): The column index letter.
    Returns:
        The column index as an integer.
    """
    letter = letter.upper()
    result = 0

    for index, char in enumerate(reversed(letter)):
        # Get the ASCII number of the letter and subtract 64 so that A
        # corresponds to 1.
        num = ord(char) - 64

        # Multiply the number with 26 to the power of `index` to get the correct
        # value of the letter based on it's index in the string.
        final_num = (26 ** index) * num

        result += final_num

    # Subtract 1 from the result to make it zero-based before returning.
    return result - 1


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description='Lonely Planet Report Order Helper')

    parser.add_argument(
        'infile',
        type=str,
        help='Infile report to read (must be xlsx)')

    parser.add_argument('-v',
                        '--verbose',
                        action='store_true',
                        help='Turn on verbose mode.')


    args = parser.parse_args()


    wb = load_workbook(filename = args.infile, use_iterators=True)
    sheet = wb.active
    #row_count = sheet.get_highest_row() - 1
    #column_count = letter_to_index(sheet.get_highest_column()) + 1
    #print row_count, column_count
    
    row_count = sheet.max_row
    column_count = sheet.max_column
    print row_count, column_count
    for i in xrange(3,row_count,1):
        isbn = "C" + str(i)
        print sheet[isbn].value
    # In [11]: wb.get_sheet_names()
    # Out[11]: ['ARK_PCA_ANT_SOLGT_OG_BEH (97)']
    # can choose active or with this name

    #for line in sheet_ranges:
    #    print line
    #print(sheet_ranges['A1'].value)
    #for i in xrange(
