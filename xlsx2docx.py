#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from optparse import OptionParser

from abstract_generator.AbstractGenerator import AbstractGenerator


def main():
    usage = 'Usage: %prog [options] input.xlsx output.docx'
    version = '%prog 20161121'
    parser = OptionParser(usage=usage, version=version)
    parser.add_option('-i', '--imagedir', metavar='IMAGE_DIR',
                      dest='image_dir', default='image',
                      help='image directory (default image)')
    parser.add_option('-t', '--template', metavar='TEMPLATE_TYPE',
                      dest='template_type', default='aini2016',
                      help='template type: \'aini2016\' or \'jscpb2016\' (default aini2016)')
    (options, args) = parser.parse_args()
    if len(args) != 2:
        parser.error('incorrect number of arguments')
    input_xlsx = args[0]
    output_docx = args[1]
    template_docx = 'template-' + options.template_type + '.docx'

    abgen = AbstractGenerator(options.image_dir, options.template_type)
    abgen.read_xlsx(input_xlsx)
    abgen.write_docx(output_docx, template_docx)
    sys.exit()


if __name__ == '__main__':
    main()
