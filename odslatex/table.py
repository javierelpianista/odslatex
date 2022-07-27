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

import numpy as np
from .style import Style
from lxml import etree
from zipfile import ZipFile
import os
import csv

class Table:
    def __init__(self,h,w):
        '''
        Create a Table object.

        Paramters:
        ----------

        h: height of the table
        w: width of the table
        '''

        self.data = []
        for _ in range(h):
            self.data.append(w*[''])

        self.borders_top   = np.zeros([h+1,w],dtype=bool)
        self.borders_left  = np.zeros([h,w+1],dtype=bool)
        self.merged = np.zeros([h,w], dtype=bool)
        self.owner  = np.zeros([h,w,2], dtype=int)
        self.sizes  = np.ones([h,w,2], dtype=int)
        self.text_alignments = np.tile('default', [h,w])

        for y in range(h):
            for x in range(w):
                self.owner[y,x,0] = y
                self.owner[y,x,1] = x

        self.h = h
        self.w = w

    def __repr__(self):
        # First we determine the max width of each column
        maxw = self.w*[0]

        # First pass. Only count width of umerged cells.
        for y in range(self.h):
            for x in range(self.w):
                #if self.merged[y][x] and not np.all(self.owner[y,x,:] == np.array([y,x])):
                if self.sizes[y][x][1] > 1: 
                    continue

                if len(self.data[y][x]) > maxw[x]:
                    maxw[x] = len(self.data[y][x])

        # Now see if the allocated space is enough for merged cells
        for y in range(self.h):
            for x in range(self.w):
                cell_width = self.sizes[y][x][1]
                if cell_width > 1:
                    cell_length = len(self.data[y][x])
                    combined_length = sum(maxw[x:x+cell_width]) + 2*(cell_width-1)

                    if cell_length > combined_length:
                        dif = cell_length - combined_length
                        extra_space = dif//cell_width
                        extra_space_2 = dif - cell_width*extra_space
                        if extra_space_2: extra_space += 1

                        for n in range(cell_width):
                            maxw[x+n] += extra_space

        ans = ''
        for y in range(self.h):
            # First draw the top border
            border_str = ' '
            for x in range(self.w):
                border_str += ''.join((maxw[x]+2)*['-' if self.borders_top[y][x] else ' '])
                border_str += ' '

            ans += border_str + '\n'

            # Now write the data
            for x in range(self.w):
                if self.sizes[y,x,1] != 0:
                    cell_width = sum(maxw[x:x+self.sizes[y,x,1]]) + 3*self.sizes[y,x,1]-1
                    fmt_str = ('|' if self.borders_left[y][x] else ' ') + '{:^' + str(cell_width) + '}' 
                    ans += fmt_str.format(self.data[y][x])

                if x == self.w-1:
                    ans += '|' if self.borders_left[y][-1] else ' '

            ans += '\n'

        border_str = ' '
        for x in range(self.w):
            border_str += ''.join((maxw[x]+2)*['-' if self.borders_top[self.h][x] else ' '])
            border_str += ' '
        ans += border_str + '\n'
        return ans

    def add_column(self, pos):
        for y in range(self.h):
            self.data[y].insert(pos, '')

        self.borders_left  = np.insert(self.borders_left, pos, self.borders_left[:,pos], axis=1)
        self.borders_top   = np.insert(self.borders_top, pos, self.borders_top[:,pos], axis=1)

        self.w += 1

    def add_row(self, pos):
        self.data.insert(pos, ['']*self.w)

        self.borders_left = np.insert(self.borders_left, pos, self.borders_left[pos,:], axis=0)
        self.borders_top = np.insert(self.borders_top, pos, self.borders_top[pos,:], axis=0)

        self.h += 1

    def merge_cells(self,y0,x0,h,w):
        self.sizes[y0,x0,:] = [h,w]

        if h == 1 and w == 1: return

        for y in range(y0,y0+h):
            for x in range(x0,x0+w):
                self.merged[y,x] = True
                self.owner[y,x,:] = [y0,x0]

                if x != x0:
                    self.borders_left[y][x] = False

                if y != y0:
                    self.borders_top[y][x] = False

                if x != x0 or y != y0:
                    self.data[y][x] = '*'
                    self.sizes[y,x,:] = 0

    def set(self,y,x,value):
        self.data[y][x] = value

    def get_cell_dimensions(self, y0, x0):
        '''
        Return the dimensions of the cell that starts in (y0,x0).
        '''

        if np.any(self.owner[y0,x0,:] != np.array([y0,x0])):
            raise Exception('The set of coordinates provided do not ' +
                    'correspond to the beginning of a cell.')

        y = y0
        x = x0

        while y < self.h and np.all(self.owner[y,x0,:] == np.array([y0,x0])):
            y+=1

        while x < self.w and np.all(self.owner[y0,x,:] == np.array([y0,x0])):
            x+=1

        return y-y0, x-x0


    def set_borders(self, y0, x0, borders):
        '''
        Set the borders of this cell. borders should be a list containing
        boolean values for [top, right, bottom, left] borders.
        If the borders are already set, then we do not overwrite.
        '''

        h, w = self.get_cell_dimensions(y0, x0)

        # Set the top and bottom borders
        for x in range(x0, x0+w):
            if borders[0]:
                self.borders_top[y0,x] = True
            if borders[2]:
                self.borders_top[y0+h,x] = True

        # Set the right and left borders
        for y in range(y0, y0+h):
            if borders[1]:
                self.borders_left[y,x0+w] = True
            if borders[3]:
                self.borders_left[y,x0] = True


    def all_elements(self):
        '''
        An iterator that runs over all the indices of the table in order, from
        left to right and then from top to bottom.

        Returns:
        --------
        (y,x) tuple
        '''

        y = 0
        x = -1

        while y < self.h:
            if x < self.w-1: 
                x += 1

            else:
                y += 1
                x = 0

            if y < self.h and np.all(self.owner[y,x,:] == [y,x]):
                yield y, x

    @classmethod
    def from_csv_file(cls, filename):
        lines = open(filename, 'r').readlines()
        data = []
        for line in csv.reader(lines):
            data.append(line)

        h = len(data)
        w = len(data[0])

        table = cls(h,w)
        table.data = data

        return table

    @classmethod
    def from_ods(cls, filename, **opts):
        options = {
                'sheet' : 0 
                }

        options.update(**opts)

        with ZipFile(filename, 'r') as zipobj:
            xml_content = zipobj.read('content.xml')
            #os.system('xmllint --format ' + os.path.join(tmpdir, 'content.xml') + '>' + os.path.join(tmpdir, 'content2.xml'))

        #tree = etree.parse('test2/content.xml')
        tree = etree.fromstring(xml_content)

        ns = {
                'table'  : 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
                'office' : 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
                'text'   : 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
                'style'  : 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
                'fo'     : 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'
                }

        default_style = Style({})

        # Read all the row, column, and cell styles
        column_styles = {}
        row_styles    = {}
        cell_styles   = {
                'Default' : default_style
                }

        for style in tree.iter(etree.QName(ns['style'],'style')):
            family = style.attrib[etree.QName(ns['style'],'family')]

            if family == 'table-row':
                family = style.attrib[etree.QName(ns['style'],'family')]
                row_styles[family] = None

            elif family == 'table-column':
                family = style.attrib[etree.QName(ns['style'],'family')]
                column_styles[family] = None

            elif family == 'table-cell':
                # Here we read the borders of the style

                # First, set default values
                borders = 4*[False]

                # Vertical alignment (TODO: not used yet)
                v_al = None

                # Text alignment
                t_al = None

                bt_tag = etree.QName(ns['fo'],'border-top')
                br_tag = etree.QName(ns['fo'],'border-right')
                bb_tag = etree.QName(ns['fo'],'border-bottom')
                bl_tag = etree.QName(ns['fo'],'border-left')
                b__tag = etree.QName(ns['fo'],'border')

                v_al_tag = etree.QName(ns['style'],'vertical-align')

                t_al_tag = etree.QName(ns['fo'],'text-align')

                style_name = style.attrib[etree.QName(ns['style'],'name')]

                for prop in style.iter(etree.QName(ns['style'],'table-cell-properties')):
                    if bt_tag in prop.attrib:
                        if 'solid' in prop.attrib[bt_tag]:
                            borders[0] = True

                    if br_tag in prop.attrib:
                        if 'solid' in prop.attrib[br_tag]:
                            borders[1] = True

                    if bb_tag in prop.attrib:
                        if 'solid' in prop.attrib[bb_tag]:
                            borders[2] = True

                    if bl_tag in prop.attrib:
                        if 'solid' in prop.attrib[bl_tag]:
                            borders[3] = True

                    if b__tag in prop.attrib:
                        if 'solid' in prop.attrib[b__tag]:
                            borders = [True, True, True, True]

                    if v_al_tag in prop.attrib:
                        v_al = prop.attrib[v_al_tag]


                for prop in style.iter(etree.QName(ns['style'],'paragraph-properties')):
                    if t_al_tag in prop.attrib:
                        t_al = prop.attrib[t_al_tag]

                cell_styles[style_name] = \
                        Style({ 'name'               : style_name,
                                'borders'            : borders})

                if t_al:
                    cell_styles[style_name].attribs['text-align'] = t_al

                if v_al:
                    cell_styles[style_name].attribs['vertical-align'] = v_al

        #for _, curr_style in cell_styles.items():
        #    print(curr_style.attribs['name'], curr_style.attribs['text-align'])

        # Now select the correct table.
        n = -1
        iterator = tree.iter(etree.QName(ns['table'],'table'))
        try: 
            while n != options['sheet']:
                tree = next(iterator)
                n += 1
        except StopIteration as e:
            print(80*'-')
            print()
            print('Table number {} not found in the file {}.'.format(options['sheet'], filename))
            print()
            print(80*'-')
            raise e


        # Count number of rows
        nrows = 0
        for row in tree.iter(etree.QName(ns['table'],'table-row')):
            nrep = 1
            key = etree.QName(ns['table'],'number-rows-repeated')
            if key in row.attrib:
                nrep = int(row.attrib[key])

            nrows += nrep

        # Count number of columns
        ncols = 0
        for col in tree.iter(etree.QName(ns['table'],'table-column')):
            nrep = 1
            key = etree.QName(ns['table'],'number-columns-repeated')
            if key in col.attrib:
                nrep = int(col.attrib[key])
            ncols += nrep

        table = cls(nrows,ncols)
        iterator = table.all_elements()

        # Read all the default cell style in each column
        column_default_styles = table.w*[None]

        n = 0
        for col in tree.iter(etree.QName(ns['table'],'table-column')):
            nrep = 1
            key = etree.QName(ns['table'],'number-columns-repeated')
            if key in col.attrib:
                nrep = int(col.attrib[key])

            key = etree.QName(ns['table'],'default-cell-style-name')
            for _ in range(nrep):
                column_default_styles[n] = col.attrib[key]
                n+=1

        # Read all the cells in the table
        for row in tree.iter(etree.QName(ns['table'],'table-row')):
            key = etree.QName(ns['table'],'number-rows-repeated')
            nrep_row = 1
            if key in row.attrib:
                nrep_row = int(row.attrib[key])

            for _ in range(nrep_row):
                for cell in row.iter(etree.QName(ns['table'],'table-cell')):
                    key = etree.QName(ns['table'],'number-columns-repeated')
                    nrep = 1
                    if key in cell.attrib:
                        nrep = int(cell.attrib[key])

                    for _ in range(nrep):
                        y,x = next(iterator)
                        
                        nrows_spanned = 1
                        ncols_spanned = 1

                        # Check how many columns the cell spans
                        key = etree.QName(ns['table'],'number-columns-spanned')
                        if key in cell.attrib:
                            ncols_spanned = int(cell.attrib[key])

                        # Now check how many rows the cell spans
                        key = etree.QName(ns['table'],'number-rows-spanned')
                        if key in cell.attrib:
                            nrows_spanned = int(cell.attrib[key])

                        # Now merge cells if required
                        if nrows_spanned > 1 or ncols_spanned > 1:
                            table.merge_cells(y,x,nrows_spanned,ncols_spanned) 

                        # Read and set the text of the cell
                        found = cell.find(etree.QName(ns['text'],'p'))
                        if found is not None:
                            table.set(y,x,found.text)

                        # Read the cell style
                        key = etree.QName(ns['table'],'style-name')

                        if key in cell.attrib:
                            style_name = cell.attrib[key]
                        else:
                            style_name = column_default_styles[x]

                        # Set borders
                        borders = cell_styles[style_name].attribs['borders']
                        table.set_borders(y,x,borders)

                        # Set text alignment
                        if cell_styles[style_name].attribs['text-align'] == 'Default':
                            key = etree.QName(ns['office'],'value-type')
                            if key in cell.attrib:
                                value_type = cell.attrib[key]
                            else:
                                value_type = None

                            if value_type == 'string' or value_type == None:
                                table.text_alignments[y,x] = 'start'
                            elif value_type == 'float':
                                table.text_alignments[y,x] = 'end'
                            else:
                                raise Exception('Unknown value type: {}'.format(value_type))
                        else:
                            table.text_alignments[y,x] = cell_styles[style_name].attribs['text-align']


        return(table)


    def draw_horizontal_border(self,y):
        borders = self.borders_top[y,:]

        lines = []
        draw  = []
        n = 0

        status = borders[0]

        for x in range(self.w):
            if borders[x] == status:
                n += 1
            else:
                lines.append(n)
                draw.append(status)
                n = 1

            status = borders[x]
            if x == self.w-1:
                lines.append(n)
                draw.append(status)
                n = 1

        ans = ''
        if len(lines) == 1:
            if draw[0]:
                ans = '\\hline'
        else:
            x = 1
            for n, line in enumerate(lines):
                if draw[n]:
                    ans += '\\cline{{{:d}-{:d}}}'.format(x,x+line-1)

                x+=line

        if len(ans):
            ans += '\n'

        return ans

    def to_latex(self):
        '''
        Return a string containing the latex code to produce the table.
        '''
        # First, get the default borders for each column
        vertical_borders = []
        for x in range(self.w+1):
            n_drawn = 0
            for y in range(self.h):
                if self.borders_left[y,x]:
                    n_drawn += 1

            if n_drawn > self.w/2:
                vertical_borders.append(True)
            else:
                vertical_borders.append(False)

        # Now get the default text alignments for each column
        default_alignments = self.w*['center']
        for x in range(self.w):
            count = {
                    'start'  : 0,
                    'center' : 0,
                    'end'    : 0
                    }

            for y in range(self.h):
                if np.all(self.owner[y,x,:] == [y,x]):
                    count[self.text_alignments[y,x]] += 1

                default_alignments[x] = max(count, key=count.get)

        # Write header
        header = '\\begin{tabular}{'

        for n, border in enumerate(vertical_borders):
            if border:
                header += '|'

            if n < self.w:
                if default_alignments[n] == 'start':
                    header += 'l'
                elif default_alignments[n] == 'center':
                    header += 'c'
                elif default_alignments[n] == 'end':
                    header += 'r'
                else:
                    raise Exception('Don''t know alignment {}'.format(default_alignments[n]))

        header += '}\n'

        body = ''

        # Draw the top horizontal border
        body += self.draw_horizontal_border(0)

        for y in range(self.h):
            curr_vert_borders = vertical_borders.copy()
            for x in range(self.w):

                if self.owner[y,x,1] == x:
                    y0, x0 = self.owner[y,x]
                    h, w = self.get_cell_dimensions(y0,x0)

                    pre_str = ''
                    post_str = ''

                    # Here we produce the alignment string. It is only relevant
                    # if:
                    #
                    # a) The borders of the current cell are different from the
                    #    default ones for this column
                    # b) The alignment of the current cell is different from
                    #    the default one for this column
                    # c) The cell occupies more than one column.
                    #
                    # Borders are only drawn to the right, except for the first
                    # column, where they are also drawn to the left.

                    alignment_str = ''
                    multicol_required = False

                    # Leftmost border of the table
                    if x == 0:
                        if self.borders_left[y,0] != curr_vert_borders[0] or w>1 or self.borders_left[y,1] != curr_vert_borders[1] or self.text_alignments[y,x] != default_alignments[0]:
                            multicol_required = True
                            alignment_str += '|' if self.borders_left[y,0] else ''

                    if self.text_alignments[y,x] == 'center':
                        alignment_str += 'c'
                    elif self.text_alignments[y,x] == 'start': 
                        alignment_str += 'l'
                    else:
                        alignment_str += 'r'

                    if x==x0 and (w>1 or self.borders_left[y,x+w] != curr_vert_borders[x+w] or self.text_alignments[y,x] != default_alignments[x]) or multicol_required:
                        multicol_required = True
                        alignment_str += '|' if self.borders_left[y,x+w] else ''

                    # Now produce a multirow or multicolumn environment if 
                    # required
                    if multicol_required:
                        pre_str += '\\multicolumn{' + str(w) + '}{' + alignment_str + '}{'
                        post_str += '}'

                    if h > 1:
                        pre_str += '\\multirow{' + str(h) + '}{*}{'
                        post_str += '}'

                    text = ''
                    if y0 == y:
                        text = self.data[y][x]

                    body += pre_str + text + post_str

                    if x+w < self.w:
                        body += ' & '
                    else:
                        body += '\\\\\n'

            # Now draw horizontal lines
            body += self.draw_horizontal_border(y+1)

        epilog = '\\end{tabular}\n' 

        return [header, body, epilog]
