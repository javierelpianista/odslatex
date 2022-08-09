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
import copy

parser = argparse.ArgumentParser(description = 'odslatex: an open-source program to convert LibreOffice Calc spreadsheets into LaTeX tables.')
parser.add_argument('-l', '--list', help='List the tables included in the ods document', action='store_true')
parser.add_argument('-n', '--which', help='Choose which table from the file you want to convert (0 by default).', default=0)
#parser.add_argument('--tmp', help='Choose the temporary directory', default='/tmp')
parser.add_argument('--no-tabular', help='Returns only the table contents, without the tabular environment definitions.', action='store_false', dest='write_tabular_environment')
parser.add_argument('--minimal-latex', help='Produce a minimal LaTeX document to compile and see the table produced.', action='store_true' )
parser.add_argument('-o', '--output-file', help='Output to file.', type=argparse.FileType('w'), default=sys.stdout)
parser.add_argument('--print-debug-info', help='Print the contents of the parsed table for debugging purposes', action='store_true')

parser.add_argument('filename', help='Filename')

def list_tables(**kwargs):
    '''
    Extract the content.xml file from the original ods file and read the sheets
    in there. Print a list with al the sheet names.

    Parameters:
    -----------

    filename: name of the file to read
    '''

    args = {
            'return_list' : False,
            'filename'     : ''
            }

    args.update(kwargs)

    with ZipFile(args['filename'], 'r') as zipobj:
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

    if args['return_list']:
        return table_names

    ans = 'List of sheets from {}:\n'.format(args['filename'])
    for n, name in enumerate(table_names):
        ans += '{:4d}: {:75}\n'.format(n, name)

    return ans

def latex_document(tables):
    '''
    Produce a LaTeX document from a list of tables

    Parameters:
    -----------

    tables: a list of tables generated with the convert_table function.
    '''

    text  = '\\documentclass{article}\n'

    if not isinstance(tables, list):
        tables_list = [tables]
    else:
        tables_list = tables

    for table_text in tables_list:
        if 'multirow' in table_text:
            text += '\\usepackage{multirow}\n'
            break

    text += '\n'
    text += '\\begin{document}\n'

    for table_text in tables_list:
        text += table_text + '\\newpage'

    text += '\\end{document}\n'

    return text

def convert_table(**kwargs):
    '''
    Convert a table read from the .ods file <filename> into LaTeX.
    If the .ods file has several sheets, then choose the <which>-th sheet.

    Parameters:
    -----------

    filename (str): name of the .ods file
    which (int): which table to convert
    '''

    args = {
            'filename'                  : '',
            'which'                     : 0,
            'print_debug_info'          : False,
            'write_tabular_environment' : True
            }

    args.update(kwargs)

    table = Table.from_ods(args['filename'],sheet=int(args['which']),print_debug_info=args['print_debug_info'])
    header, body, epilog = table.to_latex()
    body = beautify_body(body)

    if args['write_tabular_environment']:
        table_text = ''.join([header, body, epilog])
    else:
        table_text = body

    return table_text

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
    args0 = parser.parse_args()

    args = copy.copy(args0)

    if args.minimal_latex and not args.write_tabular_environment:
        raise Exception('You can either ask for a minimal LaTeX document or ' +
                'for only the table contents. Not for both.')

    if args.list:
        args.output_file.write(list_tables(args))
    else:
        if args.which == 'all':
            curr_args = vars(args).copy()
            curr_args['return_count'] = True
            ntables = int(list_tables(**curr_args))

            tables_list = []
            for n in range(ntables):
                args.which = n
                tables_list.append(convert_table(**args.to_dict()))

            if args.minimal_latex:
                args.output_file.write(latex_document(tables_list))
            else:
                for table_text in tables_list:
                    args.output_file.write(table_text + '\\newpage') 

        else:
            n = int(args.which)

            if args.minimal_latex:
                args.output_file.write(latex_document(**vars(args)))
            else:
                args.output_file.write(convert_table(**vars(args)))


if __name__ == '__main__':
    main()
