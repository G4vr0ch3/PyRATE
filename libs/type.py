#!/usr/bin/python3.10


################################################################################
#                                                                              #
#                                                                              #
#                   GNU AFFERO GENERAL PUBLIC LICENSE                          #
#                       Version 3, 19 November 2007                            #
#                                                                              #
#    This programms aims at identifying file types based on magic numbers.     #
#    Copyright (C) 2022 Gavroche                                               #
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


magic_numbers = {bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]) : 'img',
                 bytes([0xFF, 0xD8, 0xFF, 0xE0]) : 'img',
                 bytes([0xFF, 0xD8, 0xFF, 0xE1]) : 'img',
                 bytes([0xFF, 0xD8, 0xFF, 0xDB]) : 'img',
                 bytes([0xFF, 0xD8, 0xFF, 0xEE]) : 'img',
                 bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) : ['doc', 'ppt', 'xls'],
                 bytes([0x7B, 0x5C, 0x72, 0x74, 0x66, 0x31]) : 'doc', # While technically rtf document.
                 bytes([0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00]) : ['docx', 'xlsx', 'pptx', 'docm', 'xlsm', 'pptm'],
                 bytes([0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x08]) : ['ott', 'odt', 'ods', 'odp'],
                 bytes([0x25, 0x50, 0x44, 0x46]) : 'pdf',
                 }


################################################################################


def identifier(path):

    mn = ''

    info("Identifying file " + path.split("/")[-1])

    try:
        with open(path, 'rb') as in_file:
            header = in_file.read(24)

        for mag in magic_numbers:
            if header.startswith(mag):
                mn = magic_numbers[mag]

        match mn:
            case 'img':
                typ = 'img'
                success('File identified as image')

            case ['doc', 'ppt', 'xls']:
                tmp = path.split(".")[-1]
                if tmp in ['doc', 'ppt', 'xls']:
                    typ = tmp
                    success('File identified as Microsoft Office document with extension .' + typ)
                else:
                    typ = None

            case ['docx', 'xlsx', 'pptx', 'docm', 'xlsm', 'pptm']:
                tmp = path.split(".")[-1]
                if tmp in ['docx', 'xlsx', 'pptx', 'docm', 'xlsm', 'pptm']:
                    typ = tmp
                    success('File identified as Microsoft Office archive with extension .' + typ)
                else:
                    typ = None

            case ['ott', 'odt', 'ods', 'odp']:
                tmp = path.split(".")[-1]
                if tmp in ['ott', 'odt', 'ods', 'odp']:
                    typ = tmp
                    success('File identified as LibreOffice archive with extension .' + typ)
                else:
                    typ = None

            case 'pdf':
                typ = 'pdf'
                success('File identified as portable file format')
                
            case 'doc':
                typ = 'doc'
                success('File identified as RDF/DOC binary with extension .' + path.split(".")[-1])

            case _:
                typ = None


        if typ is None: fail('Unsupported file format.')

        return(typ)


    except:
        fail("File identification failed")

    return('Fail')


################################################################################


if __name__ == "__main__":

    print('Please run main.py or read software documentation')

    exit()

else:
    from .prints import *
