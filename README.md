# cidrize_wacky_ranges

Contains two scripts, json_outputter.py and standardise_ip.py. 

standardise_ip is a script that outputs standardised ip addresses from an input of a whole variety of non-standardised ip addresses.

json_outputter is a script that generates a json file that contains JSON objects of the form { Institution Name, Country, Contact, IP Range}. It runs standardise_ip to generate the standardised ip range and prepares them into a JSON object. 

To generate data.json within which JSON objects of Institution Names, Countries, Contact and IP Address Range are contained, run json_outputter.py [params] > data.json 2> errors
