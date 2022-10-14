#!/usr/bin/python3.10


################################################################################
#                                                                              #
#                                                                              #
#                   GNU AFFERO GENERAL PUBLIC LICENSE                          #
#                       Version 3, 19 November 2007                            #
#                                                                              #
#    This programms aims at sanitizing image files.                            #
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
import PIL.Image as PI


from Asterix_libs.spinner import spinner


################################################################################


def ext_img(path_in, path_out, r=True):
    info("Attempting data extraction for image " + path_in.split("/")[-1])

    try:
        # Open image file
        img_in = PI.open(path_in)

        # Get image dimension (pixels x pixels)
        (l,h) = img_in.size

    except:
        fail("Could not open source file.")
        return False, ''

    with spinner(f'Sanitizing {path_in.split("/")[-1]} ...'):

        try:
            # Create new image
            img_out = PI.new(img_in.mode, (l, h))

            # Fill new image pixel by pixel
            for i in range(l):
                for j in range(h):
                    # Get pixel color
                    pix = PI.Image.getpixel(img_in, (i,j))
                    # Fill pixel
                    img_out.putpixel((i,j), pix)

            # Save image
            img_out.save(path_out)

        except:
            fail("Data extraction attempt failed.")
            return False, ''

    success('Document sanitized successfully.')

    if r: os.remove(path_in)

    return True, path_out


################################################################################


if __name__ == "__main__":
    print('Please run main.py or read software documentation')

    exit()

else:
    from Asterix_libs.prints import *
