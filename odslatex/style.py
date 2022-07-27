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

class Style:
    def __init__(self, attribs):
        '''
        Create a style from a dictionary of attributes.
        The attributes are:

        name:       Name of the style. By default it is called "Default".
        borders:    [top, right, bottom, left], ([False, False, False, False]).
                    Four boolean values telling whether the borders of the cell
                    are drawn.
        '''

        self.attribs = self.default_attributes()
        self.attribs.update(attribs)

    @staticmethod
    def default_attributes():
        return {
                'name'           : 'Default',
                'borders'        : [False, False, False, False],
                'vertical-align' : 'Default',
                'text-align'     : 'Default',
                }
        
