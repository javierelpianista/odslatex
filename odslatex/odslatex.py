#!/bin/python

'''
odslatex: an open-source converter of LibreOffice ods spreadsheets into LaTeX code.

Copyright (c) 2020-2022 Javier Garcia.

The copyrights for code used from other parties are included in
the corresponding files.

This file is part of odslatex.

odslatex is free software; you can redistribute it and/or modify
it under the terms of the GNU Lesser General Public License as published by
the Free Software Foundation, version 3.

odslatex is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU Lesser General Public License along
with odslatex; if not, write to the Free Software Foundation, Inc.,
51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'''

import odslatex.table as tab
import sys
import os
import re

def prepare_xml_file(filename):
    from zipfile import ZipFile

    # If the file is an ods file, then extract the corresponding xml file to tmp
    if filename[-4:] == '.ods':
        tmpdir = os.path.join('/tmp', str(os.getpid()))
        with ZipFile(filename, 'r') as zipobj:
            zipobj.extract('content.xml', tmpdir)

        os.system('xmllint --format ' + os.path.join(tmpdir, 'content.xml') + ' > ' + os.path.join(tmpdir, 'content2.xml'))
        filename2 = os.path.join(tmpdir, 'content2.xml') 
    else:
        filename2 = filename
    return filename2

def get_table_instructions_from_xml(lines):
    tables = []

    reading = False
    append  = False

    for line in lines:
        match_start = re.match('<table:table\s(.*?)>', line)
        match_end   = re.match('</table:table>', line)

        if match_start:
            reading = True 
            table = []

        elif match_end:
            append  = True

        if reading:
            table.append(line)

        if append:
            tables.append(table)
            append  = False
            reading = False

    return(tables)

def get_style_instructions_from_xml(lines):
    instructions = []

    reading = False
    append  = False

    for line in lines:
        match_start  = re.search('<style:style(.*)>', line)
        match_end    = re.search('</style:style(.*)>', line)
        match_inline = re.search('<style:style(.*)/>', line)

        if match_inline:
            curr_instr = [line]
            append = True

        elif match_start:
            curr_instr = [line]
            reading = True

        elif match_end:
            curr_instr.append(line) 
            reading = False
            append = True

        elif reading:
            curr_instr.append(line)

        if append:
            instructions.append(curr_instr)
            append = False
            reading = False

    return(instructions)


def minimal_latex(filename, tables):
    string = ""

    string += '\\documentclass{article}\n\\usepackage{multirow}\n\n'
    string += '\\begin{document}\n\n'
    for nt, table in enumerate(tables):
        string += table.tabularize_latex(table.latex(header = True))
        string += '\n'

    string += '\\end{document}\n'

    return(string)

def print_help():
    for line in open('README.md', 'r').readlines():
        print(line.strip())

def main():
    args = sys.argv

    fname_in  = None
    fname_out = sys.stdout

    what = 'table'
    which = 1

    if len(sys.argv) < 2:
        print_help()
        quit()

    args.pop(0)

    fname_set = False
    debug     = False
    while args:
        arg = args.pop(0)

        if not fname_set:
            if arg == '--help':
                print_help()
                quit()
            else:
                fname_in  = arg
                fname_set = True
            continue

        if arg == '--output':
            fname_out = open(args.pop(0), 'w')
        elif arg == '--list':
            what = 'list'
            which = -1
        elif arg == '--minimal-latex':
            what = 'minimal'
        elif arg == '--table':
            what = 'table'
        elif arg == '--table-contents':
            what = 'contents'
        elif arg == '--which':
            str_arg = args.pop(0)
            if str_arg == 'all': 
                which = -1
            else:
                which = int(str_arg)
        elif arg == '--print-debug-info':
            debug = True
        elif arg == '--help':
            print_help()
            quit()
        else:
            raise Exception('Argument {} not recognized.'.format(arg))

    if not fname_in:
        raise Warning('No input file specified.')

    filename = prepare_xml_file(fname_in)

    lines = []
    for line in open(filename, 'r').readlines():
        lines.append(line.strip())

    # This is a list with xml instructions regarding the styles used in the table
    style_xml_instructions = get_style_instructions_from_xml(lines)

    # This is a list with xml instructions to read the table
    table_xml_instructions = get_table_instructions_from_xml(lines)

    tables = []

    # If the user specifies a particular table, 
    if which > 0:
        table_xml_instructions = [table_xml_instructions[which - 1]]

    for instr in table_xml_instructions:
        table = tab.Table.from_xml_instructions(instr, style_xml_instructions)
        tables.append(table)

    if debug:
        for table in tables:
            table.print_debug_info()

    print()

    if what == 'minimal':
        output = minimal_latex(filename, tables)

    elif what in ['table', 'contents']:
        output = ''
        for table in tables:
            string = table.latex(header = True if what == 'table' else False)
            string = table.tabularize_latex(string)
            output += string + '\n'

    elif what == 'list':
        output = 'Available tables are:\n'
        for n, table in enumerate(tables):
            output += '  {:2}) {}\n'.format(n+1, table.name)

    fname_out.write(output.strip() + '\n')

if __name__ == '__main__':
    main()
