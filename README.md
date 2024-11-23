# VBA---File-Light-Version-
File Selection:
We created a process where the user is asked to select an Excel file using the file dialog. The file is then opened in the macro.
2. Enable Filters:
We ensured that filters were enabled in the file, allowing data to be filtered by different columns.
3. Unhide Columns A to GX:
We implemented functionality to unhide all columns from Column A to Column GX to ensure no columns are hidden during the filtering and data removal process.
4. Filter and Remove Rows Based on Column C:
We applied a filter to Column C to only keep rows where the value is "Power Services Sales".
Any rows where Column C was something other than "Power Services Sales" were deleted.
After removing the rows, the filter on Column C was cleared.
5. Filter and Remove Rows Based on Column J:
We applied a filter to Column J to remove rows where the value was "Other OEM Fleet - Non Contractual".
After removing the unwanted rows, the filter on Column J was cleared.
6. Filter and Remove Rows Based on Column B (Countries):
We applied a filter on Column B to remove rows where the country listed was one of the following:
Afghanistan
Iran
Syria
Sudan
Palestinian Territory (Gaza Strip)
After removing the rows, the filter on Column B was cleared.
7. Save the File with Updated Name:
After the filtering and row removal process, the file was saved with the updated name “Previous Name - Column B” to reflect the changes.
Key Highlights of the Code Functionality:
User Input: The user selects a binary Excel file (e.g., .xlsb).
Data Filtering and Removal:
Column C: Filters for "Power Services Sales" and deletes other rows.
Column J: Filters out rows where the value is "Other OEM Fleet - Non Contractual".
Column B: Filters out rows with countries like Afghanistan, Iran, Syria, Sudan, and Palestinian Territory (Gaza Strip).
File Saving: The processed file is saved with a new name, indicating the column being filtered.
