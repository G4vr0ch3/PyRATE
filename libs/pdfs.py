#!/usr/bin/python3.10


################################################################################
#                                                                              #
#                                                                              #
#                   GNU AFFERO GENERAL PUBLIC LICENSE                          #
#                       Version 3, 19 November 2007                            #
#                                                                              #
#    This programms aims at sanitizing .pdf documents.                         #
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


import os

from pdf2image import convert_from_path
from PIL import Image


################################################################################


def split_img(path):

    image_list = []

    images = convert_from_path(path)

    for i in range(len(images)):
        images[i].save('Outputs/page'+ str(i) +'.png', 'PNG')
        if ext_img('Outputs/page'+ str(i) +'.png', 'Outputs/out_page'+ str(i) +'.png')[0]: image_list.append('Outputs/out_page'+ str(i) +'.png')

    return image_list


################################################################################


def recompose(image_list, out_path):

    Image.open(image_list[0]).save(out_path,
                    "PDF",
                    resolution=100.0,
                    save_all=True,
                    append_images=[Image.open(image) for image in image_list[1:]]
                    )


################################################################################


def sanitiz(path):

    info('Sanitizing ' + path.split("/")[-1])

    try:
        image_list = split_img(path)
        out_path = 'Outputs/out_' + path.split("/")[-1].split(".")[0] + ".pdf"

    except:
        fail('Data extraction failed')
        return False, ''


    info('Creating new document')

    try:
        recompose(image_list, out_path)

    except:
        fail('Document recomposition failed')
        return False, ''


    for image in image_list: os.remove(image)

    success('Document sanitized successfully.')
    return True, 'Outputs/out_' + path.split("/")[-1].split(".")[0] + ".pdf"


################################################################################


if __name__ == "__main__":
    print('Please run main.py or read software documentation')

    exit()

else:
    from .prints import *
    from .imgs import *
