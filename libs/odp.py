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


import os

import xml.etree.ElementTree as ET
import PIL.Image as PI

from odf.style import MasterPage, Style, PageLayout, GraphicProperties
from odf.opendocument import OpenDocumentPresentation
from odf.text import P
from odf.draw  import Page, Frame, TextBox, Image
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


def get_texts_from_slide(slide):

    texts = []

    # Retrieve frames from slide
    f = slide.findall('.//draw:frame', xmlns)

    for frame in f:
        # Retrieve all kinds of text in frame
        p = frame.findall('.//text:p', xmlns) + frame.findall('.//text:span', xmlns)

        for ctn in p:
            txt = {
                "ctnt": ctn.text,
                "w": frame.get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}width'),
                "h": frame.get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}height'),
                "x": frame.get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}x'),
                "y": frame.get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}y')
                }

            if txt["ctnt"] is not None : texts.append(txt)

    return texts


################################################################################


def get_imgs_from_slide(slide):

    imgs = []

    # Retrieve frames from slide with image
    f = slide.findall('.//draw:frame', xmlns)
    i = [frame for frame in f if frame.find('draw:image', xmlns) is not None]

    for ctn in i:
        img = {
            "name": 'Outputs/out_' + ctn.find('.//draw:image', xmlns).get('{http://www.w3.org/1999/xlink}href').split('/')[-1],
            "w": ctn.get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}width'),
            "h": ctn.get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}height'),
            "x": ctn.get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}x'),
            "y": ctn.get('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}y')
            }
        if img["name"] is not None: imgs.append(img)

    return imgs


################################################################################


def get_content(path):

    content = []
    id, jd = 0, 0

    # Use doc as zip file
    unzip = ZipFile(path)

    # Parse xml
    root = ET.fromstring(unzip.read('content.xml'))
    body = root.find('office:body', xmlns)
    sections = body.find('office:presentation', xmlns)
    slides = sections.findall('draw:page', xmlns)

    # Get content slide by slide
    for slide in slides:

        txts = get_texts_from_slide(slide)
        imgs = get_imgs_from_slide(slide)

        sld = {"id": slides.index(slide), "txt": txts, "img": imgs}

        content.append(sld)

    return content


################################################################################


def get_imgs(path):

    img_list = []

    # Use doc as a zip file
    unzip = ZipFile(path)

    for file in unzip.namelist():
        if file.startswith('Pictures/'):

            name = file.split("/")[-1]

            try:
                # Extracting source image
                img = PI.open(BytesIO(unzip.read(file)))
                img.save('Outputs/' + name)

                # Sanitizing image
                ext_img('Outputs/' + name, 'Outputs/out_' + name)

                # Adding sanitized image to image list
                img_list.append('Outputs/out_' + name)

            except:
                warning('Image ' + name + ' not extracted. This is not fatal. However, ' + name + ' will not be present in the sanitized document.')
                if 'table' in name.lower() and name.split(".")[-1] == "svm": warning(name + " is likely a table preview which means a table from the source document might not be displayed correctly.")


    return img_list


################################################################################


def recompose(doc, content, img_list):

    # Defining style elements
    pagelayout = PageLayout(name="MyLayout")
    t_style = Style(name="MyStyle")

    # Create slides one by one
    for s in content:

        # Create slide
        slide = Page(masterpagename=MasterPage(name=str(s["id"]), pagelayoutname=pagelayout))
        doc.presentation.addElement(slide)

        # Add text sections one by one
        for txt in s["txt"]:

            # Create frame
            s_frame = Frame(width = txt["w"], height = txt["h"], x = txt["x"], y = txt["y"])
            slide.addElement(s_frame)

            # Create text box
            txtbx = TextBox()
            s_frame.addElement(txtbx)

            # Add text
            txtbx.addElement(P(text=txt["ctnt"]))

        # Add image sections one by one
        for img in s["img"]:

            if img["name"] in img_list:
                # Create frame
                s_frame = Frame(width = img["w"], height = img["h"], x = img["x"], y = img["y"])
                slide.addElement(s_frame)

                # Import image
                href = doc.addPicture(img["name"])

                # Create image
                image = Image(href = href)

                # Add image
                s_frame.addElement(image)


################################################################################


def sanitiz(path):

    info('Sanitizing ' + path.split("/")[-1])

    # Fetching data from source
    try:
        c = get_content(path)
        i = get_imgs(path)

    except:
        fail('Data extraction failed')
        return False, ''

    info('Creating new document')

    # Creating sanitized document
    try:
        # Creating new document
        doc = OpenDocumentPresentation()

        # Creating style mandatory elements
        pagelayout = PageLayout(name="MyLayout")
        doc.automaticstyles.addElement(pagelayout)

        t_style = Style(name="MyStyle")
        t_style.addElement(GraphicProperties(fillcolor="#fff"))
        doc.automaticstyles.addElement(t_style)

        # Fill document
        recompose(doc, c, i)

        # Save document
        doc.save('Outputs/out_' + path.split("/")[-1].split(".")[0] + '.odp')

    except:
        fail('Document recomposition failed')
        return False, ''


    # Remove document images from output folder
    for f in i :
        try:
            os.remove(f)
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
