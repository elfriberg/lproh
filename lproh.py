#!/usr/bin/env python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import argparse
import numpy as np
from prettytable import PrettyTable
import time

def read_complete_list(filename, filename2):
    """
    Reads the same-folder textfile and returns array for check.
    """
    complete_list = []
    old_list = []
    with open(filename, 'r') as f:
        #read_data = f.read()
        for line in f:
            data = []
            #line = f.readline()
            words = line.split(";");
            #add_isbn = str(words[0][0:2]) + str(words[0][4]) + str(words[0][6:10]) + str(words[0][
            add_isbn = ''.join(words[0].split())
            data.append(add_isbn)
            data.append(words[1][1:])
            data.append(words[2][1:])
            data.append(words[3][1:-1])
            complete_list.append(data)
    f.closed
    print 'INFO: Aktuell LP-katalog lest inn.'

    with open(filename2, 'r') as f:
        for line in f:
            data = []
            words = line.split(";");
            add_isbn = ''.join(words[0].split())
            data.append(add_isbn)
            data.append(words[1][1:])
            data.append(words[2][1:])
            data.append(words[3][1:-1])
            old_list.append(data)
    f.closed
    print 'INFO: Liste over utdaterte LP-titler lest inn.'

    return complete_list, old_list

def show_results(detected_old_books, detected_good_books, unknown_from_report):
    print '\nTABELL 1\nFant %s utdaterte (erstattet av ny) LP-titler i din rapport:' % len(detected_old_books)
    print '(Dette programmet baserer seg på old_list.txt; du bør også sjekke siste Tabell 3 der\ntitler i rapport uten treff i lister presenteres.)'
    t = PrettyTable(['ISBN', 'Navn', 'Innbinding', 'År', 'Salg totalt', 'Beholdning'])
    #print 'ISBN-13:\tNAVN:\t\t\tINNBINDING:\tÅR:\tSALG TOTALT:\tBEHOLDNING:'
    for b in detected_old_books:
        t.add_row([b[2], b[3], b[4], b[5], b[8], b[10]])
        #print '%s\t%s\t\t%s\t\t%s\t%s\t\t%s' % (b[2], b[3], b[4], b[5], b[8], b[10])
    print t
    #print detected_old_books

    print '\nTABELL 2\nFant %s aktuelle bøker fra LP-katalogen nevnt i din rapport:' % len(detected_good_books)
    print '(Disse er sortert på beholdning (økende) slik at de øverste normalt er viktigst å bestille.)'
    t_good = PrettyTable(['ISBN', 'Navn', 'Innbinding', 'År', 'Salg totalt', 'Beholdning'])
    #print detected_good_books
    #print 'ISBN-13:\tNAVN:\t\t\tINNBINDING:\tÅR:\tSALG TOTALT:\tBEHOLDNING:'
    for b in detected_good_books:
        #print '%s\t%s\t\t%s\t\t%s\t%s\t\t%s' % (b[2], b[3], b[4], b[5], b[8], b[10])
        t_good.add_row([b[2], b[3], b[4], b[5], b[8], b[10]])
    t_good.sortby = 'Beholdning'
    print t_good

    print '\nTABELL 3\nFant %s bøker i din rapport som programmet ikke finner i innlastede godkjent/forbudt-lister:' % len(unknown_from_report)
    print '(Disse er typisk ikke-LP-titler og utdaterte LP-titler som gikk OUP før 2015-09.)'
    t_unknown = PrettyTable(['ISBN', 'Navn', 'Innbinding', 'År', 'Salg totalt', 'Beholdning'])
    for unknown in unknown_from_report:
        t_unknown.add_row([unknown[2], unknown[3], unknown[4], unknown[5], unknown[8], unknown[10]])
    t_unknown.sortby = 'Beholdning'
    print t_unknown

    qa_number = len(detected_old_books) + len(detected_good_books) + len(unknown_from_report)
    return qa_number

def show_not_found(A, np_complete_list):
    year, month = time.localtime()[0:2]
    year_month = str(year) + '-' + str(month)
    cnt_missing = 0
    cnt_notpublishedyet = 0
    t_missing = PrettyTable(['ISBN', 'Navn', 'Publikasjonsdato', 'Utgave'])
    t_notpublishedyet = PrettyTable(['ISBN', 'Navn', 'Publikasjonsdato', 'Utgave'])
    A_isbns = []
    for book in A:
        book[2] = int(book[2])
        A_isbns.append(book[2])
    #print np_complete_list
    no_titles_complete = np_complete_list.shape[0]
    #print no_titles_complete
    for i in xrange(no_titles_complete):
        if int(np_complete_list[i][0]) not in A_isbns:
            if year_month >= str(np_complete_list[i][2]):
                t_missing.add_row([np_complete_list[i][0], np_complete_list[i][1], np_complete_list[i][2], np_complete_list[i][3]])
                cnt_missing += 1
            else:
                t_notpublishedyet.add_row([np_complete_list[i][0], np_complete_list[i][1], np_complete_list[i][2], np_complete_list[i][3]])
                cnt_notpublishedyet += 1
    print '\nTABELL 4\nFant %s aktuelle bøker fra LP-katalogen som mangler i din rapport:' % cnt_missing
    print '(Du bør slå opp deres ISBN manuelt for å sjekke deres antall i beholdning, evt. generere en'
    print 'rapport som går lenger tilbake i tid. Du mangler trolig enkelte av disse titlene.)'
    t_missing.sortby = 'Publikasjonsdato'
    print t_missing
    print '\nTABELL 5\nFant %s bøker fra LP-katalogen som ikke er publisert ennå [as of %s], og som mangler i din rapport:' % (cnt_notpublishedyet, year_month)
    print '(Du bør slå opp deres ISBN manuelt for å sjekke om du har bestilt disse.)'
    t_notpublishedyet.sortby = 'Publikasjonsdato'
    print t_notpublishedyet
        

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
    not_in_report = []
    unknown_from_report = []


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
    
    complete_list, old_list = read_complete_list('complete_list.txt', 'old_list.txt')
    np_complete_list = np.array(complete_list)
    #print old_list
    np_old_list = np.array(old_list)
    #print np_old_list
    #print np_old_list[:,0]

    upperleftcell = 'A3'
    lowerrightcell = 'K' + str(row_count-2)
    #A = np.array([[i.value for i in j] for j in sheet['A3':'K305']])
    A = np.array([[i.value for i in j] for j in sheet[upperleftcell:lowerrightcell]])

    # fiks OBS

    #print A.ndim
    #print type(A[0][0])
    #for row
    for book in A:
        #print type(book[2])
        #isOld(book[2],old_books)
        book[2] = int(book[2])
        #if book[2] in old_books:
        if book[2] in np_complete_list[:,0].astype(int):
            # overwrites report-title with list-title, ok
            name_index = np.where(np_complete_list[:,0]==str(book[2]))
            name_index = name_index[-1][0]
            book[3] = np_complete_list[name_index][1]
            detected_good_books.append(book)
        else:
            if str(book[2]) in np_old_list[:,0]:
                detected_old_books.append(book)
            else:
                unknown_from_report.append(book)
            

    #print detected_old_books
    qa_number = show_results(detected_old_books, detected_good_books, unknown_from_report)
    show_not_found(A, np_complete_list)

    if int(qa_number) == int(row_count - 4):
        print 'INFO: Alle %s titler i din rapport ble klassifisert og plassert i en tabell.' % str(row_count-4)
    else:
        print 'ERROR: Bare %s titler fra din rapport ble klassifisert og plassert i en tabell,\n men rapporten inneholder %s titler! Kontakt utvikleren på even@evenezer.me og legg med rapport og utskrift!' % (qa_number, (row_count - 4))

    # show titles in complete list not in report
    
