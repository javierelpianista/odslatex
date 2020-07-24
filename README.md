# odslatex
This is a Python program that converts Libreoffice tables in `.ods` format to plain text that you can paste directly in a LaTeX document.

## Requirements
`odslatex` is designed to run with any modern GNU/Linux distribution. It runs on Python 3, and uses the following Python imports:
* `os`
* `sys`
* `re`
* `count`
* `zipfile`
* `typing`
* `numpy`

It also makes use of the `xmllint` program.

## Usage
`odslatex` does not come with an installer (yet), so the best way to use it is to give execute permissions to `odslatex.py` (updating the shebang `#!/bin/python` if necessary), and creating a link named `odslatex` located in a directory included in your `$PATH` variable. Then, it can be used with:
```
odslatex [DOCUMENT] [OPTIONS]
```
Here `[DOCUMENT]` is the LibreOffice Calc document, in `.ods` format.
The available options are:
* `--help`: Displays the help section
* `--list`: Lists the available tables in the `.ods` file
* `--minimal-latex`: Asks for a minimal LaTeX document containing the table/s
* `--table` (default): Asks for a complete `tabular` environment containing the pertaining table/s
* `--table-contents`: Returns only the table contents, without the `tabular` environment definitions.
* `--which [WHICH]`: Selects which table in the document is to be converted. `[WHICH]` can be `all`, which converts all the tables contained in the document, or a number, which only converts one of the available tables. The numbers associated with each table can be obtained with the option `--list`.

If you do not know how to create symbolic links and/or add variables to your path, then you can just wait until I write the installer, or use the files as they are by calling 
```
python /path/to/odslatex.py [DOCUMENT] [OPTIONS]
```
(assuming that your default Python executable is Python 3).

### Typical use scenario
The typical use case would be to write a LaTeX document using your preferred editor, and edit the tables with LibreOffice. You could have one file for each table, or all the tables in the same file. Then when you need the table you just call `odslatex name-of-the-table.ods`, copy the result to the clipboard and paste it in your `.tex` document.
I use Vim to edit my `.tex` files, so I have written a function that calls `fzf` with the expression `*.ods`, takes the file selected by the user and calls `odslatex` on it, pasting the result directly into the `.tex` file.

### Examples
There are some examples in the `./examples` folder. If you are within the main `odslatex` folder, you can try the following:
```
mkdir -p tex
python odslatex.py examples/example.ods --minimal-latex > tex/tex.tex
cd tex
pdflatex tex
```
This will produce a `pdf` document containing the table in the example. Another possibility is
```
python odslatex.py examples/fancy_table.ods --list
```
This will return the names of the tables available in `fancy_table.ods`.
Then you can choose to convert one with
```
python odslatex.py examples/fancy_table.ods --which N
```

## What odslatex can convert
Currently, the program can convert cells spanning multiple rows and columns into their corresponding `\multicolumn` and `\multirow` LaTeX environments. It can also detect borders and mark them, and detect horizontal alignment.

## What it can not but I'd like to add in the future
It cannot detect vertical alignment, nor can it detect formatting options, like italics, bold, or font sizes.

## Contact me
You can contact me at <javier.garcia.tw@hotmail.com>. I'm glad to receive suggestions or answer any questions.
