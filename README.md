# Tecan i-control xml-2-xlsx

## About the script

This is a script that takes an XML formatted i-control output file and generates a human readable xlsx file.
It generates a worksheet for every section it encounters, which contains a table (formatted using the wells' names) followed by the parameters the XML contains, and the start and end times for the measurement.


## Usage

- Command line
  - ```python3 xmlparser.py filename.xml```
- GUI
  - Just drag and drop the xml on the script, but please note you won't read any error in this way. Try it from the command line if it doesn't work.


## Limits

Sadly i-control's XML output isn't as complete as its xls one. There are measurements that only appear when using i-control with Excel installed, like the "Gain" value: as it doesn't appear in the XML of the same measurement, the script can't transcribe it.


## To Dos

I didn't have many outputs to test it so for now it may not handle every type of measurement.
Feel free to open issues with files it can't read and I'll try to look into it.
