rbco.msexcel
============

:Author: Rafael Oliveira <rafaelbco@gmail.com>

Overview
--------

Provide functions to read, parse and convert MS Excel spreadsheets into various
data structures.

Usage
-----

To read a MS Excel file into memory call ``xls_to_excelerator_dict(filename)``. 
This will return a dict in the *excelerator_dict* format.
Functions are provided to convert between this format and the following 
formats.

excelerator_dict
----------------

A list of tuples (sheet_name, dict). dict keys are (row_num, col_num) pairs::

    [
        (
            sheet_name,
            {
                (row_num, col_num): value,
            }
        ),
    ]
    
rows_and_columns
----------------  
    
Nested dicts which keys sheet name, row number and column number::

    {
        sheet_name: {
            row_num: {
                col_num: value,
            }
        },
    }
    
matrix
------
    
A dict mapping from sheet names to matrices, i.e, a lists of lists::

    {
        sheet_name: [
            [v01, v02, v03, ...],
            [v11, v12, v13, ...],
        ]        
    }    
    
structure
---------

Perhaps the more user-friendly format: A dict mapping from
sheet names to lists. These lists contains the rows. Each row is 
represented by a dict, mapping from column names to values. Column names
are the values in the first row of the sheet::

    {
        sheet_name: [
            {
                col_name: value,
            },
        ]        
    }


This format is only useful if the first row of each sheet is actually
a header row


