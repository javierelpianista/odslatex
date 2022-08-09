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

from setuptools import setup, find_packages

setup(
        name = 'odslatex',
        version = '0.3.1',
        author = 'Javier Garcia', 
        long_description = 'No description for now',
        packages = find_packages(),
        entry_points = {
            'console_scripts' : [
                'odslatex = odslatex.main:main'
                ]
            },
        install_requires = [
            'numpy',
            'lxml'
            ]
        )
