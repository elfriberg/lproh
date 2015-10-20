#!/usr/bin/env python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import argparse
import numpy as np
from prettytable import PrettyTable

def isOld(isbn, old_books):
    #print isbn
    for book in old_books:
        if int(isbn) == book:
            print 'old'

def show_results(detected_old_books, detected_good_books):
    print 'Found %s old books:' % len(detected_old_books)
    t = PrettyTable(['ISBN', 'Navn', 'Innbinding', 'År', 'Salg totalt', 'Beholdning'])
    #print 'ISBN-13:\tNAVN:\t\t\tINNBINDING:\tÅR:\tSALG TOTALT:\tBEHOLDNING:'
    for b in detected_old_books:
        t.add_row([b[2], b[3], b[4], b[5], b[8], b[10]])
        #print '%s\t%s\t\t%s\t\t%s\t%s\t\t%s' % (b[2], b[3], b[4], b[5], b[8], b[10])
    print t
    #print detected_old_books

    print '\nFound %s good books:' % len(detected_good_books)
    t_good = PrettyTable(['ISBN', 'Navn', 'Innbinding', 'År', 'Salg totalt', 'Beholdning'])
    #print detected_good_books
    #print 'ISBN-13:\tNAVN:\t\t\tINNBINDING:\tÅR:\tSALG TOTALT:\tBEHOLDNING:'
    for b in detected_good_books:
        #print '%s\t%s\t\t%s\t\t%s\t%s\t\t%s' % (b[2], b[3], b[4], b[5], b[8], b[10])
        t_good.add_row([b[2], b[3], b[4], b[5], b[8], b[10]])
    print t_good

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


    old_books = [9781741798227]
    permitted_lp_books = [9781742200347]

    detected_old_books = []
    detected_good_books = []


    wb = load_workbook(filename = args.infile, use_iterators=True)
    sheet = wb.active
    #row_count = sheet.get_highest_row() - 1
    #column_count = letter_to_index(sheet.get_highest_column()) + 1
    #print row_count, column_count
    
    row_count = sheet.max_row
    column_count = sheet.max_column
    #print row_count, column_count
    #for i in xrange(3,row_count,1):
    #    isbn = "C" + str(i)
    #    print sheet[isbn].value
    # In [11]: wb.get_sheet_names()
    # Out[11]: ['ARK_PCA_ANT_SOLGT_OG_BEH (97)']
    # can choose active or with this name

    #for line in sheet_ranges:
    #    print line
    #print(sheet_ranges['A1'].value)
    #for i in xrange(

    A = np.array([[i.value for i in j] for j in sheet['A3':'K305']])
    #print A.ndim
    #print type(A[0][0])
    #for row
    for book in A:
        #print type(book[2])
        #isOld(book[2],old_books)
        book[2] = int(book[2])
        if book[2] in old_books:
            detected_old_books.append(book)
        elif book[2] in permitted_lp_books:
            detected_good_books.append(book)

    #print detected_old_books
    show_results(detected_old_books, detected_good_books)
