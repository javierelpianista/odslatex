#!/bin/bash

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


# This is a small script that uses zenity to open up a dialog where you select
# the .ods file you want to convert to LaTeX.
# It is very primitive and should be improved in the future.

TMPFILE="/tmp/table.$$"
FILE=`zenity --file-selection --title="Select a File"` 

case $? in
    1) zenity --notification --window-icon="info" --text="No file selected."
        exit 0;;
    -1) zenity --notification --window-icon="error" --text="Error! Aborting."
        exit 1;;
esac

odslatex --list "$FILE" > "$TMPFILE"
tail -n+2 "$TMPFILE" > "$TMPFILE"2

NSHEETS=$(cat "$TMPFILE"2 | wc -l)

while read LINE; do
    OPTIONS+=("$LINE")
done < "$TMPFILE"2

if [[ $NSHEETS == 1 ]]; then
    N=0
else
    CHOICE=$(zenity --title="Selection" --list\
        --column="Available sheets:" "${OPTIONS[@]}")
    N=$(echo $CHOICE | cut -d: -f1)
    if ! [[ "$N" == ?(-)+([0-9]) ]]; then
        echo "$N is not a number." >&2; exit 1
    fi
fi

odslatex -n$N "$FILE" | xclip -sel clip

zenity --notification --window-icon="info" \
    --text="Copied!\nLaTeX-converted sheet $N \
           from file\n$FILE copied to clipboard"

rm "$TMPFILE" "$TMPFILE"2
