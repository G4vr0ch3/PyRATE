#!/usr/bin/python3.10


################################################################################
#                                                                              #
#                                                                              #
#                   GNU AFFERO GENERAL PUBLIC LICENSE                          #
#                       Version 3, 19 November 2007                            #
#                                                                              #
#    This programms aims at parsing binary files for content.                  #
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


from io import BytesIO


################################################################################


# Define images magic numbers based on image type
jpg_magic = [ bytes([0xFF, 0xD8, 0xFF, 0xE0]), bytes([0xFF, 0xD8, 0xFF, 0xE1]), bytes([0xFF, 0xD8, 0xFF, 0xDB]), bytes([0xFF, 0xD8, 0xFF, 0xEE]) ]
png_magic = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])


################################################################################


# Parser for jpg images
def jpg_parser(bts):

    start_offsets = []
    end_offsets = []
    jpgs = []

    # Fetching start offset based on file magic number bytes
    for i in range(len(bts)):

        infor('Fetching for start offsets : ' + str(round((i/len(bts))*100, 1)) + '% done.     ')

        if bts[i:i+4] in jpg_magic:
            start_offsets.append(i)

    # Fetching end of file offset based on termination bytes "FF D9"
    for i in range(len(bts)):

        infor('Fetching for end offsets : ' + str(round((i/len(bts))*100, 1)) + '% done.     ')

        if bts[i:].startswith(bytes([0xFF, 0xD9])):
            end_offsets.append(i)

    # Checks if files were retrieved
    if len(start_offsets) > 0 and len(end_offsets) > 0:

        start_index = 0
        prev_ind = 0

        # Makes sure the first EOF offset comes after the first start of file offset
        if start_offsets[0] >= end_offsets[0]:

            end_offsets.pop(0)

        info('Retrieved ' + str(len(start_offsets)) + ' probable JPGs.                            ' )

        # Determines image EOF offset for each start of file offset
        while start_index < len(start_offsets):

            start_offset = start_offsets[start_index]

            end_index = start_index

            # Determine image EOF while keeping in mind one image file can embed others
            while end_index < len(start_offsets) - 1 and start_offsets[end_index] < end_offsets[start_index]:
                end_index += 1

            img = {'type' : 'jpg', 'id': str(start_index), 'start_offset': start_offsets[start_index], 'end_offset': end_offsets[end_index]}

            jpgs.append(img)

            # Breaking possible infinite loop
            if start_index < end_index:
                start_index = end_index

            else:
                start_index += 1


    else:
        info('No JPG image was found in document                                ')


    return jpgs


################################################################################


# Parser for png images
def png_parser(bts):

    start_offsets = []
    end_offsets = []
    pngs = []

    # Fetching start offset based on file magic number bytes
    for i in range(len(bts)):

        infor('Fetching for start offsets : ' + str(round((i/len(bts))*100, 1)) + '% done.     ')

        if bts[i:].startswith(png_magic):
            start_offsets.append(i)

    # Fetching end of file offset based on termination bytes "FF D9"
    for i in range(len(bts)):

        infor('Fetching for end offsets : ' + str(round((i/len(bts))*100, 1)) + '% done.     ')

        if bts[i:].startswith(bytes([0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82])):
            end_offsets.append(i)

    # Checks if files were retrieved
    if len(start_offsets) > 0 and len(end_offsets) > 0:

        start_index = 0
        prev_ind = 0

        # Makes sure the first EOF offset comes after the first start of file offset
        if start_offsets[0] >= end_offsets[0]:

            end_offsets.pop(0)

        info('Retrieved ' + str(len(start_offsets)) + ' probable PNGs.                            ' )

        # Determines image EOF offset for each start of file offset
        while start_index < len(start_offsets):

            start_offset = start_offsets[start_index]

            end_index = start_index

            # Determine image EOF while keeping in mind one image file can embed others
            while end_index < len(start_offsets) - 1 and start_offsets[end_index] < end_offsets[start_index]:
                end_index += 1

            img = {'type' : 'png', 'id': str(start_index), 'start_offset': start_offsets[start_index], 'end_offset': end_offsets[end_index]}

            pngs.append(img)

            # Breaking possible infinite loop
            if start_index < end_index:
                start_index = end_index

            else:
                start_index += 1


    else:
        info('No PNG image was found in document                                ')


    return pngs


################################################################################


# Sort images based on their start offset in document
def sort_imgs(imgs):

    sorted = [imgs.pop(0)]

    for img in imgs:
        rank = len(sorted) - 1

        while rank > 0 and img['start_offset'] < sorted[rank]['start_offset']:
            rank -= 1

        sorted.insert(rank, img)

    return sorted


################################################################################


# Parse images in document
def img_parser(path):

    bin, exp = [], []
    id = 0

    # Read documents as bytes
    bts = open(path, 'rb').read()

    # Fetch images by kind
    info('Fetching JPG images.')
    bin += jpg_parser(bts)

    info('Fetching PNG images.')
    bin += png_parser(bts)

    # Sort images based on there start offset in document
    bin = sort_imgs(bin)

    # Retieve images
    for image in bin:
        ostart = image['start_offset']

        # End offset + 2 EOF bytes
        oend = image['end_offset'] + 2

        # Retrieve image as bytes
        img_bts = bts[ostart:oend]

        # Create image from bytes
        pic = PI.open(BytesIO(img_bts))
        pic.save('Outputs/' + image['id'] + '.' + image['type'])

        # Image sanitization using imgs.py
        ext_img('Outputs/' + image['id'] + '.' + image['type'], 'Outputs/out_' + image['id'] + '.' + image['type'])

        info('Sanitized image ' + str(image))

        # Add image to final image list
        exp_inf = {'id': id, 'path': 'Outputs/out_' + image['id'] + '.' + image['type']}
        exp.append(exp_inf)

        id += 1

    return exp


################################################################################


if __name__ == "__main__":
    print('Please run main.py or read software documentation')

    exit()

else:
    from .prints import *
    from .imgs import *
