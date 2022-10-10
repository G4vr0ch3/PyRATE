#!/usr/bin/python3.10


################################################################################


# This software is just a wrapper for the pyexcel_ods python library.
# It aims at sanitizing .ods documents
# This library is not licensed.


################################################################################


import pyexcel_ods as ods
import json


################################################################################


def sanitiz(path):

    info('Sanitizing ' + path.split("/")[-1])

    try:
        # Reading data from source file as a json dump
        df = ods.get_data(path)
    except:
        fail('Data extraction failed')
        return False, ''

    info('Creating new document')

    try:
        # Writing data to output file
        ods.save_data('Outputs/out_' + path.split("/")[-1].split(".")[0] + ".ods", df)
    except:
        fail('Document recomposition failed')
        return False, ''

    success('Document sanitized successfully.')
    return True, 'Outputs/out_' + path.split("/")[-1]


################################################################################


if __name__ == "__main__":
    print('Please run main.py or read software documentation')
    exit()

else:
    from Asterix_libs.prints import *
