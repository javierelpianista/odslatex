#!/bin/python

import table as tab
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

if __name__ == '__main__':
    args = sys.argv

    fname_in  = None
    fname_out = sys.stdout

    what = 'table'
    which = 1

    while args:
        arg = args.pop(0)

        if arg[:1] != '-':
            fname_in = arg
        elif arg == '--file':
            fname_in = args.pop(0)
        elif arg == '--output':
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
