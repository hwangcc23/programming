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

def review_rel_note_CR():
    return

def usage():
    print("review_rel_note_CR: find need-attention CRs in an excel file of a CR list for release note")
    print("Usage: rel_note_CR [options]")
    print("       options and arguments:")
    print("       -h: show help")
    print("       -i|--input FILENAME: give the input excel file")
    print("       -o|--output FILENAME: specify the output file name (default: \"__\"##input)")
    print("       -k|--keyword FILENAME: give the keyword file")

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
    output_file = "__" + input_file 
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
            keyword = arg

    if input_file == "":
        print("Error: input file is not given")
        usage()
        sys.exit(0)

    if keyword_file == "":
        print("Error: keyword file is not given")
        usage()
        sys.exit(0)

    review_rel_note_CR()
