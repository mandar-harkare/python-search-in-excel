# Search a text in Excel

## Introduction

This project is useful the search a text in all the sheets in the excel file and copy the rows in a new excel file.

## Inputs

* Text to search
* Path of the file to search in (If the file exists in same directory, the name will auto populate once you start typing the name)
* Select all the sheets to search the text in (This lists of all the sheets in the provided file and allows you to choose multiple sheets)

## Steps

Pull down this repository and perform the following steps

```python

pip install -r requirement.txt
python3 handler.py

>>>? Enter the search keyword: text-to-be-searched
>>>? Enter the path of excel file:  /tmp/xxxxxx.xlxs
>>>? Select sheets to search in: done (3 selections)
 >>o sheet1
 >>o sheet2
 #>>o sheet3
 #>>o sheet4
 >>o sheet5
 #>>o sheet6

Searching in following sheets:
['Sheet3', 'Sheet4', 'Sheet6']
Text found in 6 rows & copied into text-to-be-searched.xlsx

```

## Demo

![Demo](./demo.gif)
