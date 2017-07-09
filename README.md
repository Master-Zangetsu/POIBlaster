# POI Blaster

## About
POIBlaster is a java library designed to be primarily used by languages without access to Apache POI, to read and write to Microsoft Excel .xls and .xlsx files.

However methods have been made available for other users to add this as a library and use it for the same purpose should they wish to.

Java users looking for more functionality should take a look at Apache POI.

## External Usage

##### ReadOperations

ReadOperations require the following parameters to be passed to the library.
1. Operation type: OP
2. File path: FILE
3. Sheets to read: SHEETS
4. Cells to read: CELLS

All of these parameters should take the format {PARAM}:{VALUE}.
Here is an example of a some values ready to be sent to the library.

``"OP:R" + "FILE:C:\\users\{USER}\Desktop\yourfilename.xlsx";``

As you can see, for read operations the correct parameter is "OP:R".
The FILE parameter needs to be the full path to the file you wish to read including its name.

You can only send one OP type and FILE per call.

The SHEETS and CELLS parameters on the other hand support delimited strings if you want to access multiple cells across multiple sheets.
For this we use a set of designated delimiters.

Sheet level delimiter = ":"

Cell and Value level delimiter = ","

For a simple read function where you are only looking at one cell in one sheet you don't have to worry about this.

`` "SHEETS:{YOUR_SHEET_NAME}" + "CELLS:{YOUR_CELL_NAME}"``

Cell names are as you would see them in Excel, so cell "A5" represents column A row 5.

However perhaps you want to access 3 cells in 1 sheet.
In which case you should do this:

``"SHEETS:{YOUR_SHEET_NAME}" + "CELLS:{FIRST_CELL_NAME},{SECOND_CELL_NAME},{THIRD_CELL_NAME}"``

The different cells are delimited by "," to keep them separate.

You can then use the sheet delimiter to specify the sheet to access.

``"SHEETS:{FIRST_SHEET_NAME}:{SECOND_SHEET_NAME}" + "CELLS:{FIRST_CELL_NAME},{SECOND_CELL_NAME},{THIRD_CELL_NAME}:{FIRST_CELL_NAME},{SECOND_CELL_NAME}"``

In this example I am using a combination of ":" and "," to specify that the system should read three specific cells from the first sheet, and two specific cells from the second sheet.

The number of sheet delimited items across the SHEETS and CELLS must match (aka the same number of ":" must be the same).

As the system is performing its actions it will print the results out for you to read.


##### WriteOperations

WriteOperations require the all the same parameters as ReadOperations plus one more.
1. Operation type: OP
2. File path: FILE
3. Sheets to read: SHEETS
4. Cells to read: CELLS
5. Cell values: VALUES

This uses all the same rules as the CELLS parameter except that they contain the text to be added into a cell.
``"CELLS:{FIRST_CELL_NAME},{SECOND_CELL_NAME},{THIRD_CELL_NAME}:{FIRST_CELL_NAME},{SECOND_CELL_NAME}" + "VALUES:{FIRST_CELL_VALUE},{SECOND_CELL_VALUE},{THIRD_CELL_VALUE}:{FIRST_CELL_VALUE},{SECOND_CELL_VALUE} "``

As above, the number of sheet delimited items across the SHEETS, CELLS, and VALUES must match.
In addition however, the number of CELL and VALUE delimited items must also match.

The writeOperation will return a true or false string to confirm if it was succesful.




