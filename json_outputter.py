#!/usr/bin/env python2

# json_outputter.py

# json-outputter, by Chuan Tan <ct538@cam.ac.uk>
#
# Copyright (C) Chuan Tan 2018
#
# This programme is free software; you may redistribute and/or modify
# it under the terms of the Apache Licence, version 2.0.

# Reads an excel file, and processes the IP addresses using standardise-ip.
# It then packages the each row into a JSON object, and ouputs a list of these JSON Objects.
# E.g. JSON Object = {
#       "Institution" : "University of Cambridge"
#       "Country"     : "UK"
#       "Contact"     : "a123@cam.ac.uk"
#       "IP-Range"    : ["123.332.23.2/24","223.225.110.23/32"]
#       }
import sys
from openpyxl import load_workbook
import standardise_ip

def run():
    _, inputf, sheetname, institution_col_id , country_col_id, contact_col_id, ip_col_id = sys.argv
    wb = load_workbook(inputf)
    sheet = wb[sheetname]
    row_ids_ip_range = standardise_ip.process(sheet,int(ip_col_id))
    
