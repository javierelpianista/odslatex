import sys
import re
import numpy as np
from typing import Dict, Union

# This class handles the typical cells from a LibreOffice table.
# Each cell has the information about its content, its alignment, its borders and 
# how many rows and columns it occupies (for the case when one or more cells are
# combined into one). 
class Cell:
    def __init__(self):
        # How many rows and columns it occupies
        self.width   = 1
        self.height  = 1

        # The value of the cell. Currently we store it as a string
        self.val     = None
        self.covered = False

        # What type is the value? This is used to set text different text properties to different
        # value types. For example, integers and floats are right-aligned, whereas text is left-
        # aligned.
        self.value_type = None

        # Each cell has a style assigned to them. These styles define the borders and alignment.
        # A style on its own does nothing, unless it is applied to the cell. Currently this is
        # done with the Table.set_style() function.
        self.style = Style.Default()
        self.style_name = None

        # Which borders are marked?
        self.borders = {
            'top'   : 'none',
            'left'  : 'none',
            'right' : 'none',
            'bottom': 'none',
            }

        # Text alignment. Only horizontal is used for now
        # TODO program vertical alignment
        self.alignment_v = 'none'
        self.alignment_h = 'none'

    def set(self, val):
        self.val = val

    def print_debug_info(self):
        print(self.borders, self.covered, self.val, self.alignment_h, self.width, self.height)

class Style:
    def __init__(self, name: str, data: Dict[str, Union[str, int]]):
        self.name    = name
        self.data    = {}
        self.borders = {}

    @classmethod
    def Default(cls):
        style = Style('Default', dict())

        style.borders = {
            'top'   : 'none',
            'left'  : 'none',
            'right' : 'none',
            'bottom': 'none',
            }

        style.data = {
                'family'      : 'table-cell', 
                'parent'      : None, 
                'alignment_h' : None,
                'alignment_v' : None
                }

        return(style)

    def set_data(self, data_dict):
        for entry, val in data_dict.items():
            if entry == 'borders':
                self.borders = val
            else:
                self.data[entry] = val


def write_cell(cell, fmt1 = '', fmt2 = '', al = ''):
    value = str(cell.val) if cell.val else ""

    if cell.width == 1 and fmt1 == '' and fmt2 == '' and al == '':
        if cell.height == 1:
            string = value
        else:
            string = "\\multirow{" + str(cell.height) + "}{*}{" + value + "}"
    else:
        if cell.height != 1:
            value = "\\multirow{" + str(cell.height) + "}{*}{" + value + "}"

        string = "\\multicolumn{" + str(cell.width) + "}{" + fmt1.strip() + al + fmt2.strip() + "}{" + value + "}"

    return(string)
            
class Table:
    def __init__(self):
        # Cells is a list of lists; cells[i][j] contains the cell belonging to row i and column j
        self.cells  = []

        self.ncols  = 0
        self.nrows  = 0

        # This variable holds information about styles as read from the xml file. 
        # Later we interpret this and set the information for each cell.
        self.styles = [Style.Default()]

        self.name = ""

    def set_nrows(self):
        self.nrows = len(self.cells)

    def set_ncols(self):
        ncols = -1
        for n in range(nrows):
            if ncols != -1 and len(self.cells[n]) != ncols:
                raise Exception('ERROR! in Table.latex(). Not all the rows have the same amount of columns')

            ncols = len(self.cells[n])

        self.ncols = ncols

    def shape(self):
        return (self.nrows, self.ncols)

    def get_style(self, name):
        for style in self.styles:
            if style.name == name:
                return(style)

        raise Exception('In Table.get_style. Style name {} not found'.format(name))

    def get_style_value(self, name : str, value):
        for style in self.styles:
            if style.name == name:
                return(style.data[value])

    # Sets the parameters of a cell according to a style
    def set_style(self, cell, style_name):
        for style in self.styles:
            if style.name == style_name:
                cell.style = style
                cell.borders = style.borders

                if style.data['alignment_h']:
                    cell.alignment_h = style.data['alignment_h'] 
                else:
                    if cell.value_type == 'float':
                        cell.alignment_h = 'end'
                    elif cell.value_type == 'string':
                        cell.alignment_h = 'start'
                    else:
                        # This includes cells that don't have information; we need to have their alignment explicit because write_cell needs it
                        cell.alignment_h = 'end'

    def set_name(self, instructions):
        for instr in instructions:
            match_name = re.search(r'table:name="(.*?)"', instr)

            if match_name:
                self.name = match_name.group(1)

    def print_debug_info(self):
        for row in self.cells:
            for cell in row:
                cell.print_debug_info()

# ==============================================================================================================
#
# Functions for getting information from the lines of an xml file
# The file needs to be processed with xmllint --format
#
# ==============================================================================================================
    def set_shape_from_xml(self, lines):
        count_columns = False
        table_ncols = 0
        table_nrows = 0

        read_cell = False
        cells = []

        table_nrows = 0

        for line in lines:
            instruction = re.match(r'<(.*)>', line).group(1)

            # Here we read the instruction to count the number of columns.
            # We do so by looking at the table:table-column entries
            if re.match(r'table:table\s', instruction):
                count_columns = True
            elif re.match(r'/table:table', instruction) and count_columns:
                count_columns = False
            elif re.match(r'table:table-column\s', instruction) and count_columns:
                data = instruction.split(" ")

                ncols = 1
                for text in data:
                    match = re.match('table:number-columns-repeated="(.*)"', text)
                    if match:
                        ncols = int(match.group(1))

                table_ncols += ncols

            # Here we count the number of rows
            if re.match(r'table:table-row\s', instruction):
                nrows = 1
                match2 = re.search(r'table:number-rows-repeated="(.*)"', instruction)
                if match2:
                    nrows = int(match2.group(1))

                table_nrows += nrows

        self.nrows = table_nrows
        self.ncols = table_ncols

# ----------------------------------------------------------------------------------------------------------------
# Get the instructions corresponding to rows from the lines of an xml file and separate them for each row
# ----------------------------------------------------------------------------------------------------------------
    def get_row_instructions_from_xml(self, lines):
        n_repeat_row  = 1

        row_instructions = []
        n_repeat_rows = []

        read_row  = False
        read_cell = False
        reading = False

        for line in lines:
            row_start  = re.match('<table:table-row(.*)>',  line)
            row_end    = re.match('</table:table-row(.*)>', line)
            row_inline = re.match('<table:table-row(.*)/>', line)
            match_n    = re.search('number-rows-repeated="(.*)"', line)

            if row_inline:
                row_instructions.append(line)
                if match_n:
                    n_repeat_row = int(match_n.group(1))
                else:
                    n_repeat_row = 1

                n_repeat_rows.append(n_repeat_row)

            elif row_start:
                reading = True
                new_row = []
                new_row.append(line)

                if match_n:
                    n_repeat_row = int(match_n.group(1))
                else:
                    n_repeat_row = 1

            elif row_end:
                new_row.append(line)
                row_instructions.append(new_row)
                reading = False
                n_repeat_rows.append(n_repeat_row)

            elif reading:
                new_row.append(line)

        return(row_instructions, n_repeat_rows)

# ----------------------------------------------------------------------------------------------------------------
# Get the instructions corresponding to cells from those separated by row with get_row_instructions_from_xml
# ----------------------------------------------------------------------------------------------------------------
    def get_cell_instructions_from_xml_row_instructions(self, row_instructions, n_repeat_rows):
        # cell_instructions[i][j][k] contains the kth instuction for the jth cell of row i
        cell_instructions = []

        append = False
        reading = False
        n_repeat_cell = 1
        for n, ins in enumerate(row_instructions):
            curr_row_instr = []

            n_rep = 1
            for row_instruction in ins:
                cell_inline = re.match('<table:(covered-)?table-cell(.*?)/>', row_instruction)
                cell_start  = re.match('<table:(covered-)?table-cell(.*?)>',  row_instruction)
                cell_end    = re.match('</table:(covered-)?table-cell(.*?)>', row_instruction)
                match_rep   = re.search('table:number-columns-repeated="(.*?)"', row_instruction)

                if match_rep:
                    n_rep = int(match_rep.group(1))

                if cell_inline:
                    new_instr = [row_instruction]
                    append = True

                elif cell_start:
                    new_instr = [row_instruction]
                    reading = True

                elif cell_end:
                    new_instr.append(row_instruction)
                    reading = False
                    append  = True

                elif reading:
                    new_instr.append(row_instruction)

                if append:
                    append = False

                    curr_row_instr.append(new_instr)

                    for k in range(n_rep - 1):
                        curr_row_instr.append([re.sub('table:number-columns-repeated="(.*?)"', '', new_instr[0])])

            for k in range(n_repeat_rows[n]):
                cell_instructions.append(curr_row_instr)

        return(cell_instructions)

# ----------------------------------------------------------------------------------------------------------------
# Read a table from the instructions collected with get_cell_instructions_from_xml_row_instructions
# ----------------------------------------------------------------------------------------------------------------
    def read_from_xml_cell_instructions(self, cell_instructions):
        read_cell = False

        rows = []
        # This variable counts how many times it has to repeat covering cells
        carry_covered = 0

        for nr in range(self.nrows):
            cells = []

            for nc in range(self.ncols):
                nrows   = 1
                ncols   = 1
                style   = None
                covered = False
                val     = ""

                cell = Cell()
                self.set_style(cell, self.default_column_styles[nc])

                if carry_covered > 0:
                    carry_covered -= 1
                    covered = True

                for instr in cell_instructions[nr][nc]:
                    match_ncols = re.search('table:number-columns-spanned="(.*?)"', instr)
                    if match_ncols:
                        ncols = int(match_ncols.group(1))

                    match_nrows = re.search('table:number-rows-spanned="(.*?)"', instr)
                    if match_nrows:
                        nrows = int(match_nrows.group(1))

                    if re.search('covered', instr):
                        covered = True
                        match_repeat = re.search('table:number-columns-repeated="(.*?)"', instr)
                        if match_repeat:
                            carry_covered = int(match_repeat.group(1)) - 1

                    match_text = re.search('text:p>(.*?)</text:p', instr)
                    if match_text:
                        val = match_text.group(1)

                    match_style = re.search('table:style-name="(.*?)"', instr)
                    if match_style:
                        cell.style_name = match_style.group(1)

                    match_type = re.search('office:value-type="(.*?)"', instr)
                    if match_type:
                        cell.value_type = match_type.group(1)

                cell.width   = ncols
                cell.height  = nrows
                cell.covered = covered
                cell.val     = val

                if cell.style_name:
                    self.set_style(cell, cell.style_name)
                else:
                    curr_style = self.default_column_styles[nc]
                    cell.style_name = curr_style
                    self.set_style(cell, curr_style)

                cells.append(cell)

            rows.append(cells)

        self.cells = rows

    def read_style_from_xml_instructions(self, instructions):
        styles = []

        for style_instr in instructions:
            style_dict = {'borders' : dict()}

            for instr in style_instr:
                match_name = re.search('style:name="(.*?)"', instr)
                if match_name:
                    name = match_name.group(1)

                match_family = re.search('style:family="(.*?)"', instr)
                if match_family:
                    style_dict['family'] = match_family.group(1)

                match_parent = re.search('style:parent-style-name="(.*?)"', instr)
                if match_parent:
                    style_dict['parent'] = match_parent.group(1)

                match_border1 = re.search('fo:border="(.*?)"', instr)
                if match_border1:
                    expr = match_border1.group(1)

                    if 'solid' in expr:
                        for border in ['top', 'left', 'bottom', 'right']:
                            style_dict['borders'][border] = 'solid'

                else:
                    for border in ['top', 'left', 'bottom', 'right']:
                        match_border = re.search('fo:border-{}="(.*?)"'.format(border), instr)
                        if match_border:
                            expr = match_border.group(1)

                            if expr == 'none':
                                style_dict['borders'][border] = 'none'
                            else:
                                data = expr.split(" ")
                                if 'solid' in data:
                                    style_dict['borders'][border] = 'solid'

                match_alignment_v = re.search('style:vertical-align="(.*?)"', instr)
                if match_alignment_v:
                    style_dict['alignment_v'] = match_alignment_v.group(1)

                match_alignment_h = re.search('fo:text-align="(.*?)"', instr)
                if match_alignment_h:
                    style_dict['alignment_h'] = match_alignment_h.group(1)


            style = Style.Default()
            style.name = name
            style.set_data(style_dict)
            styles.append(style)

        self.styles = self.styles + styles

        # Copy the unset borders from the parent styles
        for n in range(len(self.styles)):
            if self.styles[n].data['family'] == 'table-cell':
                for border in ['top', 'left', 'bottom', 'right']:
                    if not border in self.styles[n].borders:
                        self.styles[n].borders[border] = self.get_style(self.styles[n].data['parent']).borders[border]
    
    def get_default_column_styles(self, lines):
        default_column_styles = []

        for line in lines:
            is_table_column = re.match('<table:table-column(.*)', line)

            if is_table_column:
                expr = is_table_column.group(1)
                match_default_style = re.search('table:default-cell-style-name="(.*?)"', expr)
                match_repeat        = re.search('table:number-columns-repeated="(.*?)"', expr)

                if match_repeat:
                    n_rep = int(match_repeat.group(1))
                else:
                    n_rep = 1

                if match_default_style:
                    style = match_default_style.group(1)

                for n in range(n_rep):
                    default_column_styles.append(style)

        self.default_column_styles = default_column_styles

    @classmethod
    def from_xml_instructions(cls, table_instructions, style_instructions):
        # Read the xml file
        table = Table()

        # Get the shape of the table 
        table.set_shape_from_xml(table_instructions)

        # Read the different styles
        table.get_default_column_styles(table_instructions)
        #style_instr = table.read_style_instructions_from_xml(style_instructions)
        table.read_style_from_xml_instructions(style_instructions)

        # Read the different cells
        table.set_name(table_instructions)
        a, b  = table.get_row_instructions_from_xml(table_instructions)
        cell_instr = table.get_cell_instructions_from_xml_row_instructions(a, b)
        table.read_from_xml_cell_instructions(cell_instr)

        return(table)

# ----------------------------------------------------------------------------------------------------------------
# These functions are used to convert the table into latex format
# ----------------------------------------------------------------------------------------------------------------

    # This splits multirow cells into single row cells; the one at the top contains the text while those at the bottom are empty, multicolumn rows 
    def split_vertical(self):
        nrows, ncols = self.shape()

        for nr in range(nrows):
            for nc in range(ncols):
                curr_cell = self.cells[nr][nc]

                if curr_cell.covered:
                    continue

                height = curr_cell.height
                width  = curr_cell.width

                borders = curr_cell.style.borders
                alignment_h = curr_cell.alignment_h
                alignment_v = curr_cell.alignment_v

                if height > 1:
                    # Set the bottom border to none
                    self.cells[nr][nc].borders['bottom'] = 'none'

                    cell = Cell()
                    cell.height      = height-1
                    cell.width       = width
                    cell.covered     = False
                    cell.val         = ""
                    cell.borders     = borders
                    cell.alignment_h = alignment_h
                    cell.alignment_v = alignment_v

                    # The top border must be free since this is a multirow cell
                    cell.borders['top'] = 'none'

                    self.cells[nr+1][nc] = cell

    def get_horizontal_lines(self):
        # For each row, check if the bottom or top of the cell is marked
        drawn_indices = []

        nrows, ncols = self.shape()

        for nr in range(-1, nrows):
            x = 0
            drawn_indices_row = set()

            for nc in range(ncols):
                x+=1

                if nr >= 0: 
                    cell_curr = self.cells[nr][nc]

                    if not cell_curr.covered:
                        if cell_curr.borders['bottom'] == 'solid':
                            for d in range(cell_curr.width):
                                drawn_indices_row.add(d+x)

                if nr < nrows - 1:
                    cell_next = self.cells[nr+1][nc]
                    if not cell_next.covered:
                        if cell_next.borders['top'] == 'solid':
                            for d in range(cell_next.width):
                                drawn_indices_row.add(d+x)

            drawn_indices.append(drawn_indices_row)

        # Now compact consecutive indices into tuples
        ndd = []
        for di in drawn_indices:
            nd = []

            di_l = list(sorted(di))

            if len(di_l) == 0:
                ndd.append([])
                continue

            start = di_l[0]
            new   = di_l[0]

            nd = []

            if len(di_l) == 1:
                nd = [(di_l[0], di_l[0])]

            else:
                for n in range(1, len(di_l)):
                    old = new
                    new = di_l[n]

                    if new > old + 1:
                        nd.append((start, old))
                        start = new

                if start != new: 
                    nd.append((start, new))

            ndd.append(nd)

        return(ndd)

# ----------------------------------------------------------------------------------------------------------------

    def get_vertical_lines(self):
        # For each row, check if the left or right border of the cell is marked
        drawn_indices = []

        nrows, ncols = self.shape()

        for nc in range(-1, ncols):
            x = 0
            drawn_indices_col = set()

            for nr in range(nrows):
                x+=1

                if nc >= 0: 
                    cell_curr = self.cells[nr][nc]

                    if not cell_curr.covered:
                        if cell_curr.borders['right'] == 'solid':
                            for d in range(cell_curr.height):
                                drawn_indices_col.add(d+x)

                if nc < ncols - 1:
                    cell_next = self.cells[nr][nc+1]
                    if not cell_next.covered:
                        if cell_next.borders['left'] == 'solid':
                            for d in range(cell_next.height):
                                drawn_indices_col.add(d+x)

            drawn_indices.append(sorted(drawn_indices_col))

        return(drawn_indices)

# ----------------------------------------------------------------------------------------------------------------

    def get_column_alignments(self):
        from collections import Counter

        nrows, ncols = self.shape()

        alignments = []

        for nc in range(ncols):
            col_alignments = []

            for nr in range(nrows):
                col_alignments.append(self.cells[nr][nc].alignment_h)

            count = Counter(col_alignments)
            col_alignment = count.most_common(1)[0][0]

            alignments.append(col_alignment)

        return(alignments)

# ----------------------------------------------------------------------------------------------------------------

    def alignment_char(self, alignment):
        if alignment == 'start':
            return('l')
        elif alignment == 'center':
            return('c')
        elif alignment == 'end':
            return('r')
        else:
            raise Exception('No alignment character for alignment {}'.format(alignment))

# ----------------------------------------------------------------------------------------------------------------

    def latex(self, **kwargs):
        string = ''
        nrows, ncols = self.shape()

        self.split_vertical()

        hl = self.get_horizontal_lines()
        vl = self.get_vertical_lines()

        marked_columns = []
        for item in vl:
            n_marked = len(item)
            if n_marked > nrows/2.:
                marked_columns.append(True)
            else:
                marked_columns.append(False)

        mark = marked_columns[0]

        upper_line = False

        for db in hl[0]:
            string += '\\cline{{{}-{}}}'.format(db[0], db[1])
            upper_line = True

        if upper_line: string += '\n'

        col_align = self.get_column_alignments()

        for nr in range(nrows):
            for nc in range(ncols):
                curr_cell = self.cells[nr][nc]

                if not curr_cell.covered:
                    # ---------------------------------------------------------------------------------------
                    # Check if we need to use multicolumn for the particular cell. This will happen if:
                    # a) The cell has different alignment than the column
                    # b) The cell occupies two or more spaces 
                    # c) The cell has different borders than the column 
                    use_multicolumn = False
                    fmt1 = ''
                    fmt2 = ''
                    al   = ''

                    # Check if we need to specify left border
                    if nc == 0:
                        mark_l = True if nr+1 in vl[0] else False

                        if mark_l != marked_columns[0]: use_multicolumn = True
                            #fmt1 = '|' if mark else ' '

                    if nc < ncols:
                        # Check if we need to specify right border
                        mark_r = True if nr+1 in vl[nc+1] else False

                        if mark_r != marked_columns[nc+1]: use_multicolumn = True
                            #fmt2 = '|' if mark else ' '

                    if curr_cell.alignment_h != col_align[nc] or curr_cell.width > 1: use_multicolumn = True

                    # ---------------------------------------------------------------------------------------
                    # Now define the format and alignment chars. These will be used in write_latex to 
                    # properly set up the multicolumn environment if needed.
                    if use_multicolumn:
                        fmt1 = '|' if mark_l else ' '
                        fmt2 = '|' if mark_r else ' '
                        al = self.alignment_char(curr_cell.alignment_h)

                    string += write_cell(curr_cell, fmt1, fmt2, al)

                    # Add the separator or newline symbols
                    if nc + curr_cell.width < ncols:
                        string += " & "
                    else:
                        string += "\\\\\n"

            lower_line = False
            for db in hl[nr+1]:
                string += '\\cline{{{}-{}}}'.format(db[0], db[1])
                lower_line = True

            if lower_line: string += '\n'

        if 'header' in kwargs:
            header = kwargs['header']
        else:
            header = True

        if header:
            string2 = '\\begin{tabular}{' 
            string2 += '|' if mark else ''

            for n, mark in enumerate(marked_columns[1:]):
                if   col_align[n] == 'start' : char = 'l'
                elif col_align[n] == 'center': char = 'c'
                elif col_align[n] == 'end'   : char = 'r'

                string2 += self.alignment_char(col_align[n])
                string2 += '|' if mark else ''

            string2 += '}\n'

            string = string2 + string
            string += '\\end{tabular}\n'

        return(string)

# ------------------------------------------------------------------------------
    def get_latex_column_widths(self, string):
        lines = string.split('\n')
        lines = [line.strip() for line in lines]

        widths = []

        for line in lines:
            match = re.match('(.*?)\\\\$', line)

            if match:
                info = match.group(1)
                data = info.split('&')

                lengths = []
                for word in data:
                    match_mc = re.search('\\\\multicolumn{(.*?)}', word)
                    if match_mc:
                        n = int(match_mc.group(1))
                    else:
                        n = 1

                    width = len(word.strip())
                    curr_width = int(np.ceil(width/float(n)))
                    for k in range(n):
                        lengths.append(curr_width)

                widths.append(lengths)

        widths = np.array(widths, dtype = int)

        nrows, ncols = widths.shape

        max_widths = []
        for nc in range(ncols):
            max_widths.append(max(widths[:, nc]))

        return(max_widths)

    def tabularize_latex(self, string):
        nrows, ncols = self.shape()
        widths = self.get_latex_column_widths(string)

        lines = string.split('\n')
        lines = [line.strip() for line in lines]

        string2 = ""
        for nl, line in enumerate(lines):
            match = re.match('^(.*?)\\\\\\\\$', line)

            if match:
                info = match.group(1)
                data = info.split('&')

                lengths = []

                ni = 0
                for nw, word in enumerate(data):
                    match_mc = re.search('\\\\multicolumn{(.*?)}', word)

                    if match_mc:
                        n = int(match_mc.group(1))
                    else:
                        n = 1

                    separator = ' & '
                    n1 = (n-1)*len(separator)

                    for k in range(n):
                        n1 += widths[k+ni] 

                    ni += n

                    string2 += ('{:' + str(n1) + '}').format(word.strip())

                    if nw == len(data) - 1:
                        string2 += '\\\\'
                    else:
                        string2 += separator
            else:
                string2 += line

            if nl != len(lines) - 1:
                string2 += '\n'

        return(string2)
