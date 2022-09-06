# Pyrate documentation

## Installation

The source of the software is hosted on github and should be cloned as follows:
```bash
~$ git clone https://github.com/G4vr0ch3/Pyrate
```
Before using the software, requirements have to be installed using:

```bash
~$ cd Pyrate
~/Pyrate $ pip install -r requirements.txt
```

The software is now ready to be used.

## Usage

Pyrate can be used directly from the command line or as a python library.
The “—test” flag will test the different functions of the software.

```bash
~$ ./pyrate.py --test
```

The “-f” flag with an input-file path will attempt a data sanitization process on the document.

```bash
~$ ./pyrate.py -f <INPUT_FILE>
```

The “-r” flag will remove the input file after treatment. The “—help” flag or “-h” flag will display the help page of the software.

Pyrate can also be used as a python library using the “treat” function to process documents.

```python
import pyrate

pyrate.treat("input_file_path")
```

The “test” function will test the different functions of the software.

## Supported file formats

In its current state the software supports the following file signatures:
•	Images in PNG and JPG (JPG, JPEG) formats.
•	Microsoft Office Word documents in .doc, .docx and .docm formats.
•	Microsoft Office Excel spreadsheets in .xls, .xlsx and .xlsm formats.
•	Microsoft Office PowerPoint presentations in .pptx and .pptm formats.
•	LibreOffice Writer documents in .odt and .ott formats.
•	LibreOffice Calc spreadsheets in .ods format.
•	LibreOffice Impress presentations in .odp format.
•	Portable Document Format (PDF) in .pdf format.

Some formats will never be supported for sanitizing due to the complexity of the process such as:
•	Executable files (.exe, binaries)
•	Dynamic Link Libraries (.dll)
•	Any malicious script in any programming language (.ps1, .py, .sh, .pl, .rb, .php, etc.)
Should one of these types of files be identified as malicious, they will remain untouched and will not be copied to the output medium.

Some files may also be flagged as malicious because of their content while not being dangerous as is (e.g. Text documents, XML documents, etc.). To protect ingenuous users, these files will not be copied to the output medium.

Any other file type is not supported (and files of said type will not be copied to the output medium).

## Software overview

When the Frontend is asked to sanitize files, it receives a list of the files. Pyrate takes a file path as an argument and will attempt to sanitize one file at a time. It will first identify the file’s format by comparing its magic number to a dictionary of magic numbers. If the file’s type is not part of the supported files, the software will stop trying to sanitize the file and return the information. If the format is recognized, the program will call the sanitizing function from the appropriate library to clean the file.

| ![pyrate.png](https://raw.githubusercontent.com/G4vr0ch3/Pyrate/main/pic/pyrate.png) |
|:--:|
| <b>Software architecture</b>|



Pyrate.py: This is the main program, it receives the list of files to sanitize, identifies them and uses the proper library to deal with the threat. It returns a document with the sanitizing process results.

Type.py: This is a library that aims at checking whether the file’s type is supported by the main program. Rather than just checking the file’s extension, it identifies the file’s magic number and compares it with a dictionary of magic numbers. The function used to perform this analysis is called “identifier” and returns a file type as a string.

Imgs.py: Even though the attack vectors using images are few, we decided to create a small library to handle images. It relies on the Pillow Image Library (PIL) to handle images. It is made of one function that takes an input file path and an output path as arguments. It retrieves each pixel’s color from the input file and creates a file with the same pixel colors. It returns “True” and the output’s path if the treatment succeeded and «False» and an empty file path if it did not. The sanitized file will be in the “Outputs” folder.
The treatment will preserve the image resolution as well as graphical content.

Common_bin.py: This is a library that aims at hosting common functions to parse binary files and retrieve content. It possesses multiple parsing functions that will analyze the bytes of the binary. It will identify the existence and locations of images using magic numbers and end of file (EOF) termination bytes. It then relies on PIL to convert the bytes to an image file. Each image file is then treated by the “imgs.py” library. It creates files in the “Outputs” folder. The parsing functions take paths as arguments and will return lists containing the retrieved data in the appropriate form (paths, bytes, …).

Docs.py: This is a library that aims at sanitizing Microsoft Office Word documents with the “.doc” extension. This type of word document is actually a single binary file containing the document’s content. The main sanitizing function takes a file path as an argument and first deconstructs before building a new “.docx” document. To extract the text from the document, it relies on the Antiword software [ANTIW] which is Open-Source software released under a GPL License by Adri van Os. This also provides the global layout of the document. To retrieve images from the documents, the software relies on the “common_bin.py” library’s functions. To create the new document, the software relies on the “pi-docx” library. The main sanitizing function will return “True” and the output’s path if the treatment succeeded and «False» and an empty file path if it did not. The sanitized file will be in the “Outputs” folder.
The treatment will preserve text paragraphs, images, lists, text areas content and tables.

Docxs.py: This is a library that aims at sanitizing Microsoft Office Word documents with the “.docx” or “.docm” extension. The latter file type can embed macros which, when activated, can execute malicious code on the victim’s terminal. This kind of document is an archive containing a variety of files dissecting its layout and content. The main sanitizing function takes a file path as an argument and proceeds to first extracting the relevant data from the archive before building a new “.docx” document. To extract the text and layout of the document, it relies on functions that parse the “document.xml” file contained in the archive. The different nodes of the file will provide the various section types and content. To gather the images embedded in the document, it extracts the content from the media folder of the archive. The images are then processed by the “imgs.py” library. To create the new document, the software relies on the “pi-docx” library. The main sanitizing function will return “True” and the output’s path if the treatment succeeded and «False» and an empty file path if it did not. The sanitized file will be in the “Outputs” folder.
The treatment will preserve text paragraphs, images, list items, text areas content and tables.

Odt.py: This is a library that aims at sanitizing LibreOffice Writer documents with the “.odt” or “.ott” extension. This type of document was a commonly used infection vector back in 2019 because they were not analyzed properly by anti-viruses. Like “.docx” documents, these documents are archives containing files which will dissect the documents layout and content. The main sanitizing function takes a file path as an argument and first extracts the relevant data before building a new “.odt” document. To extract the text and layout of the document, it relies on functions that parse the “content.xml” file contained in the archive. The different nodes of the file will provide the various section types and content. To gather the images embedded in the document, it extracts the content from the “Pictures” folder of the archive. The images are then processed by the “imgs.py” library. To create the new document, the software relies on the “odfpy” library. The main sanitizing function will return “True” and the output’s path if the treatment succeeded and «False» and an empty file path if it did not. The sanitized file will be in the “Outputs” folder.
The treatment will preserve text paragraphs, images, lists, text areas content and tables.


Ods.py: This is a library that aims at sanitizing LibreOffice Calc documents with the “.ods” extension. The threats concerning this file format are the same as those concerning LibreOffice Writer documents. The software is a wrapper for the “pyexcel_ods” python library. The main sanitizing function takes a file path as an argument, extracts data from the input file as a JSON dump before exporting it to a new “.ods” document.  It will return “True” and the output’s path if the treatment succeeded and «False» and an empty file path if it did not. The sanitized file will be in the “Outputs” folder. “ods.py” is not licensed.
The treatment will preserve pages count, pages names and cell values. Diagrams will be lost.

Odp.py: This is a library that aims at sanitizing LibreOffice Impress documents with the “.odp” extension. The threats concerning this file format are the same as those concerning LibreOffice Writer documents. The main sanitizing function takes a file path as an argument and first extracts the relevant data before building a new “.odp” document. To extract the text and layout of the document, it relies on functions that parse the “content.xml” file contained in the archive. The different nodes of the file will provide the various section types and content. To gather the images embedded in the document, it extracts the content from the “Pictures” folder of the archive. The images are then processed by the “imgs.py” library. To create the new document, the software relies on the “odfpy” library. The main sanitizing function will return “True” and the output’s path if the treatment succeeded and «False» and an empty file path if it did not. The sanitized file will be in the “Outputs” folder.
The treatment will preserve text paragraphs, images, text areas and table content.

Pdfs.py: This is a library that aims at sanitizing documents with the “.pdf” extension. This type of file is a common infection vector as well as a standard for document transfers. This format offers many features that can be used in malicious ways. The danger does not lie in the content of the document but in the possibility of hijacking the interpretation made by the PDF reader software. Like “.doc” documents, “.pdf” documents are not easily readable or editable. To work around this difficulty, the software first exports each page of the pdf as an image file (that are then processed by “imgs.py”) before creating a new “.pdf” document from the images. The main sanitizing function takes a file path as an argument. To convert the PDF’s pages to images, the software relies on the “pdf2image” python library. It also relies on the “PIL” library to handle images. It will return “True” and the output’s path if the treatment succeeded and «False» and an empty file path if it did not. The sanitized document will be in the “Outputs” folder.
The treatment will preserve the complete appearance of the document.

Pptxs.py: This is a library that aims at sanitizing Microsoft Office Powerpoint documents with the “.pptx” or “.pptm” extension.  The threats concerning this type of files are the same as those concerning Microsoft Office Word documents. This software relies on the “python-pptx” library to first extract the layout of the presentation before reconstructing a sanitized document. It will extract texts, images and tables. The extracted images will be treated by the “imgs.py” library. The main sanitizing function takes a file path as an argument and will return “True” and the output’s path if the treatment succeeded and «False» and an empty file path if it did not. The sanitized file will be in the “Ouptus” folder.
The treatment will preserve text paragraphs, images, list items, text areas content and tables.

Xls.py: This is a library that aims at sanitizing Microsoft Office Excel documents with the “.xls”, “.xlsx” or “.xlsm” extension. The latter file type can embed macros which, when activated, can execute malicious code on the victim’s terminal. The main sanitizing function takes a file path as an argument and proceeds to first extracting the relevant data before building a new “.xlsx” document. It relies on the “xlrd” library to extract pages and cell values. To create the new document, the software makes use of the “xlswriter” library. The main sanitizing function will return “True” and the output’s path if the treatment succeeded and «False» and an empty file path if it did not. The sanitized file will be in the “Outputs” folder.
The treatment will preserve pages count, pages names and cell values. Diagrams will be lost.

Prints.py: This is a library that aims at printing information regarding the program’s execution using the following color code:
-	Green: success
-	Blue: information
-	Orange: warnings
-	Red: failures
“prints.py” is not licensed.

Hash.py: This is a library that aims at helping in checking data integrity. It relies on the “hashlib” python library and the “sha512” (Secure Hash Algorithm-512) cryptographic hash algorithm. The main function takes a file path as an argument and returns the sha512-hash for the file.
“hash.py” is not licensed.
 
