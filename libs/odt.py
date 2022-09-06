#!/usr/bin/python3.10


################################################################################
#                                                                              #
#                                                                              #
#                   GNU AFFERO GENERAL PUBLIC LICENSE                          #
#                       Version 3, 19 November 2007                            #
#                                                                              #
#    This programms aims at sanitizing .odt and .ott documents.                #
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


import xml.etree.ElementTree as ET
import PIL.Image as PI

from os import remove
from odf.opendocument import OpenDocumentText, OpenDocumentDrawing
from odf.draw import Frame, Image, Page, Object
from odf.style import Style, GraphicProperties
from odf.text import P, List, ListItem, ListStyle
from odf.table import Table, TableRow, TableColumn, TableCell
from odf import teletype
from zipfile import ZipFile
from io import BytesIO


################################################################################


# Define XML namespace (LibreOffice)
xmlns = {
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
        'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
        'svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
        'xlink': 'http://www.w3.org/1999/xlink',
        }


################################################################################


# Get content from a section
def get_content(section, id):

    ctnt = []

    # Get all texts from the section.
    raw_texts = section.text

    span_texts = section.findall('text:span', xmlns)

    p_texts = section.findall('.//text:p', xmlns)
    p_span_texts = [s.findall('text:span', xmlns) for s in p_texts]

    # Get image node from the section. None type if none
    image = section.find('.//draw:image', xmlns)

    # Add different text types accordingly.
    if raw_texts is not None:
        ctnt.append({'type': 'txt', 'ctnt': raw_texts})

    if p_span_texts != []:
        # List items
        ctnt.append({'type': 'txt', 'ctnt': ' '.join([t.text for ps in p_span_texts for t in ps])})

    elif span_texts != []:
        # Other text forms
        ctnt.append({'type': 'txt', 'ctnt': ' '.join([t.text for t in span_texts])})

    else:
        # Text paragraphs
        ctnt.append({'type': 'txt', 'ctnt': ' '.join([t.text for t in p_texts])})


    if image is not None:
        # Gets image id if it exists
        ctnt.append({
                    'type': 'img',
                    'file': 'Outputs/out_' + section.find('.//draw:image', xmlns).get('{http://www.w3.org/1999/xlink}href').split('/')[-1],
                    'w': section.find('.//draw:frame', xmlns).get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}width') ,
                    'h': section.find('.//draw:frame', xmlns).get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}height')
                    })
        id += 1

    return ctnt, id


################################################################################


def get_sections(path):

    content = []
    id, jd = 0, 0

    # Use doc as zip file
    unzip = ZipFile(path)

    # Parse xml
    root = ET.fromstring(unzip.read('content.xml'))
    body = root.find('office:body', xmlns)
    sections = body.find('office:text', xmlns)

    # Retrieve p sections content
    for index in range(len(sections)):

        section = sections[index]
        sec_type = section.tag.split("}")[-1]

        if sec_type == 'p':

            ctnt, id = get_content(section, id)

            if ctnt != []: content += ctnt

        elif sec_type == 'table':

            cells = {'type': 'table', 'id': jd, 'dim': (0,0), 'cells': []}

            # Gets table data if it exists
            rows = section.findall('.//table:table-row', xmlns)

            # Retrieve cells
            for row in rows:
                columns = row.findall('.//table:table-cell', xmlns)
                for column in columns:
                    for section in column.findall('text:p', xmlns):
                        ctnt, id = get_content(section, id)
                        if ctnt != []:

                            cell = (tuple(ctnt), (rows.index(row), columns.index(column)))

                            cells['cells'].append(cell)

            # Add table dimensions
            dim = (len(rows), len(columns))
            cells['dim'] = dim

            content.append(cells)
            jd += 1

        elif sec_type == 'list':
            lst_items = {'type': 'list', 'items': []}

            for item in section.findall('text:list-item', xmlns):

                # Get list item
                ctnt, id = get_content(item.find('text:p', xmlns), id)

                lst_items['items'].append(ctnt)

            # Add lst to content
            content.append(lst_items)

        else:
            warning('Unrecognized section. This is not fatal.')




    return content


################################################################################


def get_imgs(path):

    img_list = []

    # Use doc as a zip file
    unzip = ZipFile(path)

    for file in unzip.namelist():
        if file.startswith('Pictures/'):

            # Extracting source image
            name = file.split("/")[-1]
            img = PI.open(BytesIO(unzip.read(file)))
            img.save('Outputs/' + name)

            # Sanitizing image
            ext_img('Outputs/' + name, 'Outputs/out_' + name)

            # Adding sanitized image to image list
            img_list.append('Outputs/out_' + name)


    return img_list


################################################################################


def gen_drw(item):

    # Create temporary drawing document
    draw_doc = OpenDocumentDrawing()
    page = Page(masterpagename = u"new")
    draw_doc.drawing.addElement(page)

    # Create picture frame
    frame = Frame(layer = 'layout', width=item['w'], height=item['h'])
    page.addElement(frame)

    # Add picture
    href = draw_doc.addPicture(item['file'])
    img = Image(href = href)
    frame.addElement(img)

    # Return picture
    return draw_doc


################################################################################


def get_elements(cells, r, c):

    elem_lst = []

    # Retrieve element for cell in position (r,c)
    for cell in cells:

        if (r,c) == cell[1]:
            elem_lst.append(cell)

    return elem_lst


################################################################################


def recompose(doc, content, img_list):

    # Document reconstitution
    for item in content:

        if item['type'] == 'txt':
            doc.text.addElement(P(text=item['ctnt']))

        elif item['type'] == 'img':

            # Create new paragraph
            p = P()

            # Create frame
            doc.text.addElement(p)
            df = Frame(width=item['w'], height=item['h'], anchortype='as-char')
            p.addElement(df)

            # Get image
            drw = gen_drw(item)

            # Add image
            loc = doc.addObject(drw)
            fi = Object(href = loc)
            df.addElement(fi)

        elif item['type'] == 'list':

            # Create list (bullet list by default)
            lst = List(stylename="BulletList")

            # Add list items
            for bull in item['items']:

                # Only supports text items
                if bull != [] and bull[0]['type'] == 'txt':
                    itm = ListItem()
                    itm_para = P()
                    itm_ctnt = bull[0]['ctnt']
                    teletype.addTextToElement(itm_para, itm_ctnt)
                    itm.addElement(itm_para)
                    lst.addElement(itm)

            # Add list to document
            doc.text.addElement(lst)

        elif item['type'] == 'table':

            # Table creation
            table = Table(name='Tbl' + str(item['id']))

            # Retrieve table dimensions
            rows,cols = item['dim']

            # Add columns to table
            table.addElement(TableColumn(numbercolumnsrepeated=cols))

            # Fill table by rows
            for row in range(rows):

                # Create new row
                r = TableRow()
                table.addElement(r)

                # Fill cells
                for col in range(cols):

                    # Create cell
                    c = TableCell()
                    r.addElement(c)

                    # Retrieve cell content
                    elements = get_elements(item['cells'], row, col)

                    # Fill cell with content
                    for element in elements:

                        for ctnt in element[0]:

                            if ctnt['type'] == 'txt':

                                p = P(text=str(ctnt['ctnt']))
                                c.addElement(p)

                            elif ctnt['type'] == 'img':

                                # Create new paragraph
                                p = P()

                                # Create frame
                                c.addElement(p)
                                df = Frame(width=ctnt['w'], height=ctnt['h'], anchortype='as-char')
                                p.addElement(df)

                                # Get image
                                drw = gen_drw(ctnt)

                                # Add image
                                loc = doc.addObject(drw)
                                fi = Object(href = loc)
                                df.addElement(fi)

            # Add table to document
            doc.text.addElement(table)


################################################################################


def sanitiz(path):

    info('Sanitizing ' + path.split("/")[-1])

    # Fetching data from source
    try:
        c = get_sections(path)
        i = get_imgs(path)

    except:
        fail('Data extraction failed')
        return False, ''


    info('Creating new document')

    # Creating sanitized document
    try:

        # New document creation
        doc = OpenDocumentText()
        recompose(doc, c, i)
        doc.save('Outputs/out_' + path.split("/")[-1])

    except:
        fail('Document recomposition failed')
        return False, ''

    # Remove document images from output folder
    for image in i:
        try:
            os.remove(image)
        except:
            warning('Cleaning failed for image ' + image)

    success('Document sanitized successfully.')
    return True, 'Outputs/out_' + path.split("/")[-1]


################################################################################


if __name__ == "__main__":
    print('Please run main.py or read software documentation')

    exit()

else:
    from .imgs import *
    from .prints import *
