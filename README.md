# odslatex
This is a Python program that converts Libreoffice tables in `.ods` format to plain text that you can paste directly in a LaTeX document.

## Requirements
`odslatex` is designed to run with any modern GNU/Linux distribution. It runs on Python 3, and depends on `numpy` and `lxml`.

## Usage
`odslatex` comes with a `setup.py` script. The easiest way to install it is `cd`-ing to the project folder and running 
```
pip install .
```

Once installed, it can be used with:

```
odslatex [DOCUMENT] [OPTIONS]
```
Here `[DOCUMENT]` is the LibreOffice Calc document, in `.ods` format.
The available options are:
* `-h,--help`: Displays the help section
* `-l,--list`: Lists the available tables in the `.ods` file
* `-n,--which [WHICH]`: Selects which table in the document is to be converted. `[WHICH]` can be `all`, which converts all the tables contained in the document, or a number, which only converts one of the available tables. The numbers associated with each table can be obtained with the option `--list`. By default it is equal to 0.
* `-o,--output-file`: Output to a file instead of the standard output.
* `--minimal-latex`: Asks for a minimal LaTeX document containing the selected table. This document can be readily compiled to see if the table looks like it should.
* `--no-tabular`: Returns only the table contents, without the `tabular` environment definitions.
* `--print-debug-info`: Prints the contents, borders and alignment of every cell in every table.

### Typical use scenario
The typical use case would be to write a LaTeX document using your preferred editor, and edit the tables with LibreOffice. You could have one file for each table, or all the tables in the same file. Then when you need the table you just call `odslatex name-of-the-table.ods`, copy the result to the clipboard and paste it in your `.tex` document.

### Examples
There are some examples in the `./examples` folder. If you are within the main `odslatex` folder, you can try the following:
```
mkdir -p tex
odslatex examples/fancy.ods --minimal-latex --output-file tex/example.tex
cd tex
pdflatex example
```
This will produce a `pdf` document containing the table in the example. Another possibility is
```
odslatex examples/fancy_table.ods --list
```
This will return the names of the tables available in `fancy_table.ods`.
Then you can choose to convert one with
```
odslatex examples/fancy_table.ods --which N
```

### Graphical interface? 
It would be nice to have a graphical interface. For now, the `odslatex_zenity.sh` script is provided. It opens a file-select prompt where you can choose the file you want to convert. If the selected Calc document has more than one sheet, it asks you to select which sheet to convert. The result is pasted into the clipboard. To run it, you need to have installed both `zenity` and `xclip`, and some notification daemon running. Most Linux distributions should have these by default.

## What odslatex can convert
Currently, the program can convert cells spanning multiple rows and columns into their corresponding `\multicolumn` and `\multirow` LaTeX environments. It can also detect borders and mark them, and detect horizontal alignment.

## What it can not but I'd like to add in the future
It cannot detect vertical alignment, nor can it detect formatting options, like italics, bold, or font sizes.

## Contact me
You can contact me at <javier.garcia.tw@hotmail.com>. I'm glad to receive suggestions or answer any questions.
