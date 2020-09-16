# Dependencies

## Python
You need to have Python 3.6+ installed on your system but to have good performace I recommend to use latest stable version of python (3.8.*).

If you are Linux or Mac user python installation can be done using package manager. If you are a Windows user visit
[Python Downloads](https://www.python.org/downloads/)

## OpenPyXL

You need to have `OpenPyXL` library installed.
To installed it execute following command on the terminal or powershell.

```bash
pip install openpyxl
```

# How to use

You need to run `main.py` script using python. There are two ways to give the source file to the script.

##  As a command line argument

```bash
python main.py <source_file> <output_file>
```
output_file argument is optional. If it is not specified script will save the output file next to the source file.


## In run time

If you run the script without giving any command line arguments, script will ask about source file,

```bash
python main.py
[FILE]: Enter source file path ((relative / absolute)):
SOURCE FILE: Source.xlsx
[FILE]: Enter output file path ((relative / absolute)):
OUTPUT FILE (_output_Source.xlsx): output.xlsx
```