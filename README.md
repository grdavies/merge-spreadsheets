# Spreadsheet Merging Tool
This is a simple python tool to merge any number of spreadsheets that exist in the same directory structure. If there are sub-directories any spreadsheet under the provided directory will also be merged. If sheet names match across each sheet they will get merged into one. If one sheet has columns that are not in another, they will be added. 

The output filename is defaulted to `merged.xlsx` and it will be created in the directory provided.

## Pre-Reqs
Python 3.x 

## Installation
```
git clone https://github.com/grdavies/merge-spreadsheets
cd merge-spreadsheets
pip install -r requirements.txt
```

## Usage
Run `python main.py` and enter the path when prompted.
