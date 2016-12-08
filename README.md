# abstract_generator
![python-2.7, 3.5-blue](https://img.shields.io/badge/python-2.7, 3.5-blue.svg)
![license](https://img.shields.io/badge/license-apache-blue.svg)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/DaisukeMiyamoto/abstract_generator/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/DaisukeMiyamoto/abstract_generator/?branch=master)

generate abstract document from excel table

## dependencies
- pandas
- xlrd
- python-docx
- pillow

## usage
```
Usage: xlsx2docx.py [options] input.xlsx output.docx

Options:
  --version             show program's version number and exit
  -h, --help            show this help message and exit
  -i IMAGE_DIR, --imagedir=IMAGE_DIR
                        image directory (default image)
  -t TEMPLATE_TYPE, --template=TEMPLATE_TYPE
                        template type: 'aini2016' or 'jscpb2016' (default
                        aini2016)
```
