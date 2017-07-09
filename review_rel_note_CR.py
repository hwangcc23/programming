#!/usr/bin/python
# This Python file uses the following encoding: utf-8

"""
review_rel_note_CR.py helps me automate a fucking boring stuff: find need-attention CRs
having some keywords in an excel file in which a CR list is provided for the release note
of S/W package.

Copyright (C) 2017 Chih-Chyuan Hwang (hwangcc@csie.nctu.edu.tw)

This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License version 2 as
published by the Free Software Foundation.
"""

import sys
import getopt
import openpyxl
import copy

def usage():
    print("review_rel_note_CR: find need-attention CRs in an excel file of a CR list for release note")
    print("Usage: rel_note_CR [options]")
    print("       options and arguments:")
    print("       -h: show help")
    print("       -i|--input FILENAME: give the input excel file")
    print("       -o|--output FILENAME: specify the output file name (default: \"__\"##input)")
    print("       -k|--keyword FILENAME: give the keyword file")

def get_keywords(keyword_file):
    keywords = []
    try:
        f = open(keyword_file, "r")
        try:
            while 1:
                line = f.readline()
                if line == "":
                    break
                if line[0] == "#":
                    continue
                keyword = line.strip('\n')
                if keyword == "":
                    continue
                keywords.append(copy.copy(keyword))
        except IOError:
            print("Fail to read " + keyword_file)
        finally:
            f.close()
    except IOError:
        print("Fail to open " + keyword_file)
    #print(keywords)
    return keywords

def review_rel_note_CR(input_file, output_file, keyword_file):
    print("input_file = " + input_file, ", output_file = " + output_file, ", keyword_file = " + keyword_file)
    print("Generating...")

    keywords = get_keywords(keyword_file)
    if len(keywords) == 0:
        print("No keyword is found")
        print("Abort")
        return

    try:
        wb = openpyxl.load_workbook(input_file)
    except IOError:
        print("Fail to load workbook from " + input_file)
        print("Abort")
        return

    sheet = wb.active

    Titles = []
    title_row = 1
    for row in range(1, sheet.max_row + 1):
        if sheet.cell(row=row, column=1).value == "CR ID":
            title_row = row
            break
    for col in range(1, sheet.max_column + 1):
        cell = sheet.cell(row=title_row, column=col)
        Titles.append(copy.copy(cell.value))

    CRs = []
    for row in range(title_row, sheet.max_row + 1):
        CR = {}
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=col)
            CR[sheet.cell(row=title_row, column=col).value] = cell.value
        CRs.append(copy.copy(CR))

    print("done")
    return

if __name__ == "__main__":
    argv = sys.argv[1:]
    if len(argv) == 0:
        usage()
        sys.exit(0)

    try:
        opts, args = getopt.getopt(argv, "hi:o:k:", ["input=", "output=", "keyword="])
    except getopt.GetoptError:
        usage()
        sys.exit(0)

    input_file = ""
    output_file = ""
    keyword_file = ""

    for opt, arg in opts:
        if opt == "-h":
            usage()
            sys.exit(0)
        elif opt in ("--input", "-i"):
            input_file = arg
        elif opt in ("--output", "-o"):
            output_file = arg
        elif opt in ("--keyword", "-k"):
            keyword_file = arg

    if input_file == "":
        print("Error: input file is not given")
        usage()
        sys.exit(0)

    if output_file == "":
        output_file = "__" + input_file

    if keyword_file == "":
        print("Error: keyword file is not given")
        usage()
        sys.exit(0)

    review_rel_note_CR(input_file, output_file, keyword_file)
