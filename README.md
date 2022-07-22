# Tecan i-control xml-2-xlsx

## About the script

This is a script that simply takes an XML formatted i-control output file and generates a human readable xlsx file.
It generates a worksheet for every section it encounters, which contains a table (formatted using the wells' names) followed by the parameters the XML contains, and the start and end times for the measurement.

It is not a very refined output for now, but I just needed to get it working.


## Usage

- Command line
  - ```python3 xmlparser.py filename.xml```
- GUI
  - Just drag and drop the xml on the script


## To Dos

A lot!
It still needs help polishing its output formatting, also more thorough error checking. I didn't have many outputs to test it but it should handle well any well-formed output.
