#!/bin/bash

output="test/test.tex"
mkdir -p test

echo "\\documentclass{article}" > $output
echo "\\usepackage{multirow}" >> $output
echo >> $output
echo "\\begin{document}" >> $output

names=("example" "fancy" "instruments" "problem")
for name in ${names[@]}; do
	echo "\\section*{$name}" >> $output
	python odslatex.py examples/$name.ods >> $output
	echo $\\vspace{1cm}$ >> $output
done

echo >> $output
echo "\\end{document}" >> $output

cd test
pdflatex test
cp test.pdf ..
cd ..

echo "Removing test directory..."
rm -rf test

echo 
echo "Test compiled. Check test.pdf."
