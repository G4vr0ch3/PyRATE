#!/usr/bin/python3.10


################################################################################
#                                                                              #
#                                                                              #
#                   GNU AFFERO GENERAL PUBLIC LICENSE                          #
#                       Version 3, 19 November 2007                            #
#                                                                              #
#    This programms aims at sanitizing virus-infected files.                   #
#    Copyright (C) 2022 Gavroche, Roxane                                       #
#                                                                              #
#    This program is free software: you can redistribute it and/or modify      #
#    it under the terms of the GNU Affero General Public License as published  #
#    by the Free Software Foundation, either version 3 of the License, or      #
#    (at your option) any later version.                                       #
#                                                                              #
#     This program is distributed in the hope that it will be useful,          #
#    but WITHOUT ANY WARRANTY; without even the implied warranty of            #
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the             #
#    GNU Affero General Public License for more details.                       #
#                                                                              #
#    You should have received a copy of the GNU Affero General Public License  #
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.    #
#                                                                              #
#                                                                              #
################################################################################


import argparse
import json

from pyfiglet import figlet_format as pff

from libs import docs, docxs, hash, imgs, odp, ods, odt, pdfs, pptxs, prints, type, xls
from datetime import datetime
from libs.prints import *


################################################################################


def tests():

    warning('Software test started on [' + str(datetime.now().strftime("%d/%m/%Y-%H:%M:%S")) + ']')

    #   Test examples :
    docs.sanitiz('Inputs/test.doc')
    docxs.sanitiz('Inputs/test.docx')
    try:
        success('Hash : ' + hash.sha('Inputs/test.jpg'))
    except:
        fail('Hash failed.')
    imgs.ext_img('Inputs/test.jpg', 'Outputs/out_test.jpg', False)
    odp.sanitiz('Inputs/test.odp')
    ods.sanitiz('Inputs/test.ods')
    odt.sanitiz('Inputs/test.odt')
    pdfs.sanitiz('Inputs/test.pdf')
    type.identifier('Inputs/test.jpg')
    xls.sanitiz('Inputs/test.xlsm')
    pptxs.sanitiz('Inputs/test.pptx')

    warning('Software test ended on [' + str(datetime.now().strftime("%d/%m/%Y-%H:%M:%S")) + ']')


################################################################################


def treat(f_path):

    f_type = type.identifier(f_path)

    # Attempt data extraction based on type
    match f_type:

        case "img":

            stat, o_path = imgs.ext_img(f_path, 'Outputs/out_' + f_path.split("/")[-1], False)

        case 'doc':

            stat, o_path = docs.sanitiz(f_path)

        case 'docx' | 'docm':

            stat, o_path = docxs.sanitiz(f_path)

        case 'odt' | 'ott':

            stat, o_path = odt.sanitiz(f_path)

        case 'xls' | 'xlsx' | 'xlsm':

            stat, o_path = xls.sanitiz(f_path)

        case 'pdf':

            stat, o_path = pdfs.sanitiz(f_path)

        case 'ods':

            stat, o_path = ods.sanitiz(f_path)

        case 'odp':

            stat, o_path = odp.sanitiz(f_path)

        case 'pptx' | 'pptm' | 'ppt':

            stat, o_path = pptxs.sanitiz(f_path)

        case None:

            exit()

        case _:

            fail('Something weird happened. Exiting.')
            exit()

    try:
        # Save operation results
        with open("san_results.json", "r+") as results:

            # Read existing results
            f_data = json.load(results)

            # Create JSON dump individual results
            ind_result = {
                "Date": datetime.now().strftime("%d/%m/%Y-%H:%M:%S"),
                "FileName": f_path.split("/")[-1],
                "FileType": f_type,
                "SANSTat": stat,
                "OUTPATH": o_path,
                "HASH": hash.sha(o_path),
            }

            # Append individual data
            f_data["ind_results"].append(ind_result)

            results.seek(0)

            js = json.dumps(f_data, indent=4)

            # Write to result file
            results.write(js)

    except:
        fail("Failed to write scan results. The file will be considered as not sanitized.")


    try:
        # Remove source file if required
        if args.r : os.remove(f_path)
    except:
        fail('Failed to remove source file. This is not fatal.')

    # Print completion status
    if stat:
        success('Done')
    else:
        fail('Done')


################################################################################


if __name__ == '__main__':

    # Software definition
    print(pff("Pyrate"))

    # Argument parser creation
    parser = argparse.ArgumentParser(description='PYthon Recomposition of AlTered Elements. Copyright (C) 2022 Gavroche, Roxane.')

    parser.add_argument('--test', help="Run software test", action="store_true")
    parser.add_argument('-f', metavar='file', type=str, help="Target file path")
    parser.add_argument('-r', help="Source file removal flag", action="store_true")

    args=parser.parse_args()

    if args.test: tests(); exit()

    # Retrieve file path and type
    f_path = args.f

    treat(f_path)
