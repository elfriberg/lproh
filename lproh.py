#!/usr/bin/env python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import argparse
import numpy as np
from prettytable import PrettyTable
import time
import urllib2
import os.path
import sys
import datetime
import io, json
import readline

__author__ = "Even Langfeldt Friberg"
__copyright__ = "Copyright 2015, Lonely Planet Report Order Helper"
__license__ = "GPL"
__version__ = "1.0"
__maintainer__ = "Even Langfeldt Friberg"
__email__ = "even@evenezer.me"
__status__ = "Production"

def download_lists():
    """
    Downloads current LP complete stock list and incomplete LP outdated
    stocks lists from authors' website. They are used to compare against
    stock found in imported xlsx report.
    """
    try:
        url = "http://folk.uio.no/evenlf/lp-lists/complete_list.txt"
        file_name = url.split('/')[-1]
        u = urllib2.urlopen(url)
        f = open(file_name, 'wb')
        meta = u.info()
        file_size = int(meta.getheaders("Content-Length")[0])
        print "Laster ned: %s byte: %s" % (file_name, file_size)

        file_size_dl = 0
        block_sz = 8192
        while True:
            buffer = u.read(block_sz)
            if not buffer:
                break

            file_size_dl += len(buffer)
            f.write(buffer)
            status = r"%10d  [%3.2f%%]" % (file_size_dl,
                                           file_size_dl * 100. / file_size)
            status = status + chr(8) * (len(status) + 1)
            print status,
        f.close()
    except urllib2.HTTPError, err:
        if err.code == 404:
            print 'FEIL: Filen %s ble ikke funnet. Error 404: Listen kan ha blitt flyttet, eller siden er nede.' % url
            if os.path.isfile('complete_list.txt'):
                try:
                    mtime = os.path.getmtime('complete_list.txt')
                except OSError:
                    mtime = 0
                last_modified_date = datetime.datetime.fromtimestamp(mtime)
                print 'VELG: Hvis du vil fortsette med gammel liste, trykk g etterfulgt av ENTER på tastaturet -- alt annet avslutter.'
                print 'INFO: Gammel liste ble sist endret %s' % last_modified_date
                while True:
                    choice = raw_input("> ")

                    if choice == 'g':
                        print "Bruker gammel liste."
                        break
                    else:
                        print 'Avslutter.'
                        sys.exit(1)
            else:
                print 'INFO: Filen complete-list.txt eksisterer ikke i programmets mappe. Avslutter.'
                sys.exit(1)
        else:
            raise
    except urllib2.URLError, err:
        print 'ERROR: Du er ikke koblet til Internett. Programmet avsluttes.'
        sys.exit(1)

    try:
        url = "http://folk.uio.no/evenlf/lp-lists/old_list.txt"
        file_name = url.split('/')[-1]
        u = urllib2.urlopen(url)
        f = open(file_name, 'wb')
        meta = u.info()
        file_size = int(meta.getheaders("Content-Length")[0])
        print "Laster ned: %s byte: %s\n" % (file_name, file_size)

        file_size_dl = 0
        block_sz = 8192
        while True:
            buffer = u.read(block_sz)
            if not buffer:
                break

            file_size_dl += len(buffer)
            f.write(buffer)
            status = r"%10d  [%3.2f%%]" % (file_size_dl,
                                           file_size_dl * 100. / file_size)
            status = status + chr(8) * (len(status) + 1)
            print status,
            f.close()
    except urllib2.HTTPError, err:
        if err.code == 404:
            print 'FEIL: Filen %s ble ikke funnet. Error 404: Listen kan ha blitt flyttet, eller siden er nede.' % url
            if os.path.isfile('old_list.txt'):
                try:
                    mtime = os.path.getmtime('old_list.txt')
                except OSError:
                    mtime = 0
                last_modified_date = datetime.datetime.fromtimestamp(mtime)
                print 'VELG: Hvis du vil fortsette med gammel liste, trykk g etterfulgt av ENTER på tastaturet -- alt annet avslutter.'
                print 'INFO: Gammel liste ble sist endret %s' % last_modified_date
                while True:
                    choice = raw_input("> ")

                    if choice == 'g':
                        print "Bruker gammel liste."
                        break
                    else:
                        print 'Avslutter.'
                        sys.exit(1)
            else:
                print 'INFO: Filen old_list.txt eksisterer ikke i programmets mappe. Avslutter.'
                sys.exit(1)
        else:
            raise
    except urllib2.URLError, err:
        print 'ERROR: Du er ikke koblet til Internett.'

def load_json(filename):
    with io.open('{0}.json'.format(filename), encoding='utf-8') as f: 
        return f.read()

def read_complete_list(filename):
    """
    Reads the (downloaded) same-folder stock lists.

    If ISBN is found in both complete_list and old_list, it is not accepted
    in complete_list.
    """

    completefile = load_json(filename)
    booklist = json.loads(completefile, encoding="utf-8")
    
    print 'INFO: Aktuell LP-katalog lest inn.'

    return booklist


def load_spreadsheet():
    wb = load_workbook(filename=args.infile, use_iterators=True)
    sheet = wb.active
    row_count = sheet.max_row
    column_count = sheet.max_column
    upperleftcell = 'A3'
    lowerrightcell = 'K' + str(row_count - 2)
    read_spreadsheet = np.array([[i.value for i in j] for j in sheet[upperleftcell:
                                                  lowerrightcell]])
    return read_spreadsheet, row_count


def find_active_and_replaced(read_spreadsheet, complete_list):
    processed_isbns = []
    for book in read_spreadsheet:
        rsbook_isbn = str(book[2])
        if rsbook_isbn in complete_list:
            if complete_list[rsbook_isbn]["Status"] == "active":
                report_active.append(book)
                processed_isbns.append(rsbook_isbn)
            elif complete_list[rsbook_isbn]["Status"] == "replaced":
                report_replaced.append(book)
                processed_isbns.append(rsbook_isbn)
        else:
            processed_isbns.append(rsbook_isbn)
            report_unknown.append(book)
        
    
    return report_active, report_replaced, report_unknown, processed_isbns

def find_not_present(complete_list, report_active, report_replaced, processed_isbns):
    for clbook in complete_list:
        if clbook not in processed_isbns and complete_list[clbook]["Status"] == "active":
            report_not_present.append(clbook)  
    return report_not_present


def show_tables(report_active, report_replaced, report_unknown, report_not_present, year_month, hide):
    """
    Print tables to screen.
    """
    
    # Active titles found
    print '\nTABELL 1\nFant %s aktuelle bøker fra LP-katalogen nevnt i din rapport:' % len(report_active)
    t_active = PrettyTable(['ISBN', 'Navn', 'År', 'Salg totalt', 'Beholdning'])
    for book in report_active:
        t_active.add_row([book[2], complete_list[book[2]]["Title"], book[5], book[8], book[10]])
    t_active.sortby = 'Beholdning'
    print t_active

    # Active titles missing
    print '\nTABELL 2\nDin rapport mangler %s aktuelle bøker fra LP-katalogen:' % len(report_not_present)
    t_not_present = PrettyTable(['ISBN', 'Navn', 'Publikasjonsdato', 'Utgave', 'Kommentar'])
    for book in report_not_present:
        bookdate = time.strptime(complete_list[book]["DatePublished"], "%Y-%m")
        now = time.strptime(year_month, "%Y-%m")
        if bookdate > now:
            comment = "N.Y.P."
        else:
            comment = ""
        t_not_present.add_row([book, complete_list[book]["Title"], complete_list[book]["DatePublished"], complete_list[book]["Edition"], comment])
    t_not_present.sortby = "Publikasjonsdato"
    print t_not_present

    # Replaced titles found
    print '\nTABELL 3\nFant %s utdaterte (erstattet av ny) LP-titler i din rapport:' % len(report_replaced)
    t_replaced = PrettyTable(['ISBN', 'Navn', 'År', 'Salg totalt', 'Beholdning'])
    for book in report_replaced:
        t_replaced.add_row([book[2], complete_list[book[2]]["Title"], book[5], book[8], book[10]])
    t_replaced.sortby = "Beholdning"
    print t_replaced

    # Unknown titles found
    if hide == False:
        print '\nTABELL 4\nFant %s ukjente titler i din rapport:' % len(report_unknown)
        t_unknown = PrettyTable(['ISBN', 'Navn', 'År', 'Salg totalt',
                                 'Beholdning'])
        for unknown in report_unknown:
            t_unknown.add_row([unknown[2], unknown[3], unknown[5],
                               unknown[8], unknown[10]])
        t_unknown.sortby = 'Beholdning'
        print t_unknown

    qa_number = len(report_active) + len(report_replaced) + len(report_unknown)
    return qa_number


def calculate_date():
    year, month = time.localtime()[0:2]
    if len(str(month)) != 2:
        month = '0' + str(month)
    year_month = str(year) + '-' + str(month)
    return year_month

if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description='Lonely Planet Report Order Helper')

    parser.add_argument(
        'infile',
        type=str,
        help='Infile report to read (must be xlsx)')

    parser.add_argument('-v',
                        '--verbose',
                        action='store_true',
                        help='Turn on verbose mode.')

    parser.add_argument('-hide',
                        '--hide',
                        action='store_true',
                        help='Hide unknown titles.')

    args = parser.parse_args()

    report_active = []
    report_replaced = []
    report_not_present = []
    report_unknown = []

    year_month = calculate_date()

    read_spreadsheet, row_count = load_spreadsheet()
    #download_lists()
    complete_list = read_complete_list('testdict')

    report_active, report_replaced, report_unknown, processed_isbns = find_active_and_replaced(read_spreadsheet, complete_list)
    report_not_present = find_not_present(complete_list, report_active, report_replaced, processed_isbns)

    qa_number = show_tables(report_active, report_replaced, report_unknown, report_not_present, year_month, args.hide)

    if int(qa_number) == int(row_count - 4):
        print 'INFO: Alle %s titler i din rapport ble klassifisert og plassert i en tabell.' % str(
            row_count - 4)
    else:
        print 'ERROR: Bare %s titler fra din rapport ble klassifisert og plassert i en tabell,\n men rapporten inneholder %s titler! Kontakt utvikleren på even@evenezer.me og legg med rapport og utskrift!' % (
            qa_number, (row_count - 4))
