# read-graph-from-excel

This module is a POC to read an excel spreadsheet ('xlsx') and generate a directed graph from the formulas in the spreadsheet.

The assumption is that an excel spreadsheet can be seen as one big function that maps a set of input parameters to a set of output parameters.

The idea is to visualize the dependencies between the formulas of the spreadsheed in a directed graph.

The nodes represent the cells with the formulas, e.g. node B1 has formula "B1=A1+1". The edges represent how the formulas are linked to each other, e.g. node B1 "B1=A1+1" is linked to node A1.

Nodes with no ingoing edges represent the "input parameters" of the excel spreadsheet while nodes with on outgoing edges represent the "output parameters".

## Curent state of the script

- Can read an excel spreadsheet
- Parse the cells on one Sheet until max_row and max_col
- Infer the dependencies between the cells
- Create and plot a graph that shows the nodes and the dependencies

## Todo

- Extend to read multiple sheets
- Extend to dynamically infer the range of the active cells in a sheet (max_row and max_col are currently hardcoded)
- Test with more sophisticated spreadsheets
- Refactor and clean up the code :-)

## Dependencies

- https://openpyxl.readthedocs.io
- https://github.com/xflr6/graphviz

