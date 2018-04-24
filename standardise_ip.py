# standardise_ip.py

# standardise-ip, by Chuan Tan <ct538@cam.ac.uk>
#
# Copyright (C) Chuan Tan 2018
#
# This programme is free software; you may redistribute and/or modify
# it under the terms of the Apache Licence, version 2.0.

# Reads from excel spreadsheet and converts them into IPNetwork, then outputs them
# usage: standardise_ip [inputfile] [outputfile]
# You can easily check if an IPAddress is in IPNetwork by using
# If IPAddress in IPNetwork:
#           do whatever


from cidrize import cidrize,CidrizeError
from netaddr import *
from openpyxl import load_workbook
import re
# Global variables
badones =[]; # Weeding out the bad IP addresses
goodones = []; # Just for record purposes


# Patterns
pat_privateip = re.compile("192\.168\..*\..*"); # 192.168.*.*
pat_hyphen = re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}-\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"); # match xx.xxx.xxx.xx-xx.xxx.xx.xx
pat_bracket = re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/\d{1,2}"); # (xx.xx.xx.xx/xx)
pat_wildcard = re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}\.\*"); # xx.xxx.xx.*
pat_thirdoctet_wildcard = re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}-\d{1,3}\.\*"); # xx.xxx.xx-xx.*
pat_thirdoctet = re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}-\d{1,3}\.\d{1,3}"); # xx.xxx.xx-xx.xx
pat_words = re.compile("[^0-9.-/*]"); # EZProxyIP123.405.2.3
pat_two_brackets = re.compile("\d{1,3}\.\d{1,3}\.\(\d{1,3}-\d{1,3}\)\.\(\d{1,3}-\d{1,3}\)") # xx.xxx.(xx-xx).(xx-xx)
pat_simple = re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}") # xx.xxx.xx.xx
pat_fourthoctet = re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}-\d{1,3}") # xx.xxx.xx.xx-xxx
pat_squarebrackets = re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1}\[\d{1,3}-\d{1,3}\]"); # xx.xxx.xxx.x[0-255]
pat_squarebrackets_fourth = re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}\.\[\d{1,4}]"); # xx.xxx.xxx.[xxxx]



# Dictionary of stuff to replace. For standardising.
def replace_all(text,dic):
    for i,j in dic.iteritems():
        text = text.replace(i,j)
    return text

reps = {"to":"-","and":"&", \
        u"\u002D":"-", \
        u"\u058A":"-", \
        u"\u05BE":"-", \
        u"\u1400":"-", \
        u"\u1806":"-", \
        u"\u2011":"-", \
        u"\u2012":"-", \
        u"\u2013":"-", \
        u"\u2014":"-", \
        u"\u2015":"-", \
        u"\u2E3A":"-", \
        u"\u2E3B":"-", \
        u"\uFE58":"-", \
        u"\uFE63":"-", \
        u"\uFF0D":"-", \

        }


def thirdoctet(ip_result):
    prefix = re.search(re.compile("\d{1,3}\.\d{1,3}"),ip_result).group(0)
    suffix = re.split(re.compile("\d{1,3}\.\d{1,3}\.\d{1,3}-\d{1,3}"),ip_result)[1]
    ip_range = re.search(re.compile("\d{1,3}-\d{1,3}"),ip_result).group(0)
    first = ip_range.split("-")[0]
    second = ip_range.split("-")[1]

    if suffix == ".*":
        startip = prefix + "." + first  + ".0"
        endip = prefix + "." + second + ".255"
    else:
        startip = prefix + "." + first  + suffix
        endip = prefix + "." + second  + suffix

    result_list = iprange_to_cidrs(startip,endip)
    return result_list

def two_brackets(ip_result):
    prefix = re.search(re.compile("\d{1,3}\.\d{1,3}"),ip_result).group(0)
    third_octet = re.findall("\d{1,3}-\d{1,3}",ip_result)[0]
    fourth_octet = re.findall("\d{1,3}-\d{1,3}",ip_result)[1]

    third_octet_first = third_octet.split("-")[0]
    third_octet_second = third_octet.split("-")[1]

    fourth_octet_first = fourth_octet.split("-")[0]
    fourth_octet_second = fourth_octet.split("-")[1]


    startip = prefix + "." + third_octet_first + "." + fourth_octet_first
    endip = prefix + "." + third_octet_second + "." + fourth_octet_second

    result_list = iprange_to_cidrs(startip,endip)
    return result_list

def categorise(ip): # categorise individual ips
    ip = ip.replace(" ","")
    if re.search(pat_privateip,ip) is None:
            try:

                if re.search(pat_hyphen,ip) is not None:
                    result = re.search(pat_hyphen,ip).group(0)
                    ip_split = result.split("-")
                    startip = ip_split[0]
                    endip = ip_split[1]
                    cidrizedarray = iprange_to_cidrs(startip,endip)
                    return cidrizedarray

                elif re.search(pat_bracket,ip) is not None:
                    ip_result = re.search(pat_bracket,ip).group(0)
                    return cidrize(ip_result)

                elif re.search(pat_thirdoctet_wildcard,ip) is not None:
                    ip_result = re.search(pat_thirdoctet_wildcard,ip).group(0)
                    cidrizedarray = thirdoctet(ip_result)
                    return cidrizedarray

                elif re.search(pat_thirdoctet,ip) is not None:
                    ip_result = re.search(pat_thirdoctet,ip).group(0)
                    cidrizedarray = thirdoctet(ip_result)
                    return cidrizedarray

                elif re.search(pat_wildcard,ip) is not None:
                    ip_result = re.search(pat_wildcard,ip).group(0)
                    cidrizedarray = cidrize(ip_result)
                    return cidrizedarray


                elif re.search(pat_two_brackets,ip) is not None:
                    cidrizedarray = two_brackets(re.search(pat_two_brackets,ip).group(0))
                    return cidrizedarray

                elif re.search(pat_fourthoctet,ip) is not None:
                    ip_result = re.search(pat_fourthoctet,ip).group(0)
                    cidrizedarray = cidrize(ip_result)
                    return cidrizedarray

                elif re.search(pat_squarebrackets,ip) is not None:
                    ip_result = re.search(pat_squarebrackets,ip).group(0)
                    cidrizedarray = cidrize(ip_result)
                    return cidrizedarray

                elif re.search(pat_squarebrackets_fourth,ip) is not None:
                    ip_result = re.search(pat_squarebrackets_fourth,ip).group(0)
                    cidrizedarray = cidrize(ip_result)
                    return cidrizedarray

                elif re.search(pat_simple,ip) is not None:
                    ip_result = re.search(pat_simple,ip).group(0)
                    cidrizedarray = cidrize(ip_result)
                    return cidrizedarray

                else:
                    # doesn't fit into any defined form. Must be a typo somewhere.
                    return "fail"
            except CidrizeError as e:
                return "fail"
    return None


def screen(rawvalue): # digest the rawvalue of row

    rawvalue = replace_all(rawvalue,reps)
    ss = [x.strip() for x in re.split('[,;&]',rawvalue)]; # Remove whitespace and produce an array of (hopefully) readable IP Addresses.
    cidrizedarray = []
    for ip in ss:
        if ip is None or ip is u'':
            continue
        cidrizedip = categorise(ip)

        # failsafe
        if cidrizedip == "fail":
            #cidrizedip = human_input(ip,rawvalue)
            badones.append(ip)

        if type(cidrizedip) == list:
            for i in cidrizedip:
                if i == None or i is u'':
                    continue
                cidrizedarray.append(i)
        elif cidrizedip == None or cidrizedip is u'' :
            continue
        else:
            cidrizedarray.append(cidrizedip)
    return cidrizedarray

def human_input(ip,rawvalue):
    # Last resort, ask a user to verify ip address and correct it.
    # Returns none if ip can't be recognised by user, the cidrized ip if ip is corrected.
    print "I can't process %s. \n It's from %s \n Please correct it. (type i to ignore): " %(ip,rawvalue)
    rawinput = raw_input("")
    if rawinput == "i":
        return None
    rawinput = replace_all(rawinput,reps)
    ss = [x.strip() for x in re.split('[,;&]',rawinput)]
    for ip in ss:
        if ip is None or ip is u'':
            continue
        cidrizedip = categorise(ip)
        if cidrizedip == "fail":
            cidrizedip = human_input(ip,rawinput)
        else:
            return cidrizedip




    return network


def process(sheet): # process each row
    for i in range(1,sheet.max_row+1):
        if(sheet.cell(column = 1,row = i).value) is not None:
            rawvalue = sheet.cell(column=1,row=i).value
            cidrized = screen(rawvalue)
            print "Processed row %d" %(i)
            sheet.cell(column =2,row=i).value = str(cidrized)
            goodones.append(cidrized)

    for ip in badones:
        print ip

    print "There are %d bad ips need fixing." %(len(badones))


def alternate_way_of_running():
    wb = load_workbook("ip-ranges.xlsx")
    sheet = wb.get_sheet_by_name("Sheet1")
    process(sheet)
    wb.save("ip-ranges-edited.xlsx")

# Preprocessing
def run():
    _, inputf, outputf = sys.argv
    wb = load_workbook(inputf)
    sheet = wb["Sheet1"]
    process(sheet)

if __name__ == "__main__":
    run()
