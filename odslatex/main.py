# odslatex: an open-source program to convert LibreOffice Calc spreadsheets 
# into LaTeX tables.
#
# Copyright (c) 2022, Javier Garcia.
#
# This file is part of odslatex.
#
# odslatex is free software: you can redistribute it and/or modify it under the
# terms of the GNU General Public License as published by the Free Software
# Foundation, either version 3 of the License, or (at your option) any later
# version.
#
# odslatex is distributed in the hope that it will be useful, but WITHOUT ANY
# WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR
# A PARTICULAR PURPOSE. See the GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License along with
# odslatex. If not, see <https://www.gnu.org/licenses/>.

import argparse
from zipfile import ZipFile
from lxml import etree
import sys
from .table import Table

parser = argparse.ArgumentParser(description = 'Hello')
parser.add_argument('-l', '--list', help='List the tables included in the ods document', action='store_true')
parser.add_argument('-n', '--which', help='Choose which table from the file you want to convert (0 by default).', default=0)
#parser.add_argument('--tmp', help='Choose the temporary directory', default='/tmp')
parser.add_argument('--table-contents', help='Returns only the table contents, without the tabular environment definitions.', action='store_true')
parser.add_argument('--minimal-latex', help='Produce a minimal LaTeX document to compile and see the table produced.', action='store_true')
parser.add_argument('-o', '--output-file', help='Output to file.', type=argparse.FileType('w'), default=sys.stdout)
parser.add_argument('--print-debug-info', help='Print the contents of the parsed table for debugging purposes', action='store_true')

parser.add_argument('filename', help='Filename')

def list_tables(args):
    '''
    Extract the content.xml file from the original ods file and read the sheets
    in there. Print a list with al the sheet names.

    Parameters:
    -----------

    filename: name of the file to read
    '''
    with ZipFile(args.filename, 'r') as zipobj:
        xml_content = zipobj.read('content.xml')

    tree = etree.fromstring(xml_content)

    ns = {
            'table'  : 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
            'office' : 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
            'text'   : 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
            'style'  : 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
            'fo'     : 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'
            }

    table_names = []
    for table in tree.iter( etree.QName(ns['table'],'table') ):
        table_names.append(
                table.attrib[ etree.QName(ns['table'],'name') ]
                    )

    print( 'List of sheets from {}:'.format(args.filename) )
    for n, name in enumerate(table_names):
        print( '{:4d}: {:75}'.format(n, name) )

def convert_table(args):
    '''
    Convert a table read from the .ods file <filename> into LaTeX.
    If the .ods file has several sheets, then choose the <which>-th sheet.

    Parameters:
    -----------

    filename (str): name of the .ods file
    which (int): which table to convert
    '''

    table = Table.from_ods(args.filename,sheet=int(args.which))
    header, body, epilog = table.to_latex()
    body = beautify_body(body)

    table_text = ''.join([header, body, epilog])

    text = table_text
    if args.minimal_latex:
        text  = '\\documentclass{article}\n'
        if 'multirow' in table_text:
            text += '\\usepackage{multirow}\n'

        text += '\n'
        text += '\\begin{document}\n'
        text += table_text
        text += '\\end{document}\n'
    return text

def beautify_body(body):
    import re 

    lines = body.split('\n')
    data = [list(map(lambda x: x.strip(), line.split('&'))) for line in lines]

    # Maximum number of columns (there should be at least one line without 
    # merged columns)
    maxcols = max(list(map(len, data)))

    lines_dict = {
            'merged' : [],
            'full'   : [],
            'single' : []
            }

    kind_list = []

    # Separate lines according to the number of columns they have
    for elem in data:
        if any(map(lambda x: 'multicolumn' in x, elem)) and len(elem) < maxcols:
            lines_dict['merged'].append(elem)
            kind_list.append('merged')
        elif len(elem) == maxcols:
            lines_dict['full'].append(elem)
            kind_list.append('full')
        elif len(elem) == 1:
            lines_dict['single'].append(elem)
            kind_list.append('single')
        else:
            raise Exception('Something''s wrong with line ', elem)

    # Now determine the maximum length of each column in the full lines (i.e.),
    # lines that contain the maximum number of columns (because they don't 
    # have multicolumn environments)
    lens = []
    for elem in lines_dict['full']:
        lens.append(list(map(len, elem)))

    lens2 = list(map(list, zip(*lens)))

    # Determine the maximum length of each column (only full columns)
    maxlens = list(map(max, lens2))

    # Now see if the compacted columns can be fit into those lengths. If not,
    # update them. We make the rightmost element grow until it fits.
    merged_fmt_strs = []
    ncols_list = []

    for elem in lines_dict['merged']:

        x = 0
        curr_ncols_list = []
        for n, cell in enumerate(elem):
            m = re.search(r'\\multicolumn\{(.*?)\}', cell)
            ncols = 1
            if m:
                ncols = int(m.group(1))

            curr_ncols_list.append(ncols)

            maxlen = sum(maxlens[x:x+ncols])
            dif = len(cell) - maxlen
            if dif > 0:
                maxlens[x+ncols-1] += dif

            x += ncols

        ncols_list.append(curr_ncols_list)

    # With the updated maximum lengths on each column, prepare a list with the
    # format strings for the lines with merged columns
    for n, elem in enumerate(lines_dict['merged']):
        curr_fmt_str = ''
        x = 0
        for m, cell in enumerate(elem):
            ncols = ncols_list[n][m]
            maxlen = sum(maxlens[x:x+ncols]) + (ncols-1)*3 # Add space for ' & '

            curr_fmt_str += '{{:{}{:d}}}'.format('^' if ncols > 1 else '', max(maxlen, len(cell)))

            if m != len(elem)-1:
                curr_fmt_str += ' & '

            x += ncols

        merged_fmt_strs.append(curr_fmt_str)

    fmt_str = ''
    for n, l in enumerate(maxlens):
        fmt_str += '{{:{:d}}}'.format(l)
        if n != len(maxlens)-1:
            fmt_str += ' & '
    
    body = ''
    for kind in kind_list:
        elem = lines_dict[kind].pop(0)
        if kind == 'full':
            body += fmt_str.format(*elem) + '\n'
        elif kind == 'single':
            body += elem[0] + '\n'
        elif kind == 'merged':
            body += merged_fmt_strs.pop(0).format(*elem) + '\n'

    return body

def main():
    args = parser.parse_args()

    if args.list:
        list_tables(args)
    else:
        args.output_file.write(convert_table(args))

if __name__ == '__main__':
    main()
