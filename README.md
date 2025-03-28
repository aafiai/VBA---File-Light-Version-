# VBA---File-Light-Version-


*The VBA macro script is designed to process a selected `.xlsb` file by applying various filters to its worksheet and then saving the modified version of the file with a new name. Here's a step-by-step breakdown of its functionality:*

 

1. **File Selection**: Prompts the user to select a `.xlsb` file using a file dialog. If no file is selected, the macro exits.

 

2. **Worksheet Definition**: Assumes the first sheet in the workbook is the target worksheet (`ws`).

 

3. **Enable Filters**: Ensures filters are enabled on the first row of the worksheet.

 

4. **Unhide Columns**: Unhides all columns from A to GX.

 

5. **Column C Filtering**: Filters out rows based on specific sales categories ("Steam Power Sales", "No GE internal Sales Channel", "Oil & Gas Sales") and deletes visible rows.

 

6. **Column J Filtering**: Filters for "Other OEM Fleet - Non Contractual" and deletes visible rows.

 

7. **Column B Filtering**: Filters rows based on specific countries and deletes visible rows.

 

8. **Column G Filtering**: Filters rows based on the status ("Retired", "Inactive", "Never Built (Cancelled)", "Scrapped") and deletes visible rows.

 

9. **Column H Filtering**: Filters for plant types ("Steam Plant", "Nuclear", "Aero Plant", "Other Plant", "Industrial") and deletes visible rows.

 

10. **Column Y Filtering**: Filters for specific equipment and deletes visible rows.

 

11. **Column Z Filtering**: Filters based on status similar to Column G and deletes visible rows.

 

12. **Column AM Filtering**: Filters specific OEM categories ("Gas Aero OEM", "Gas Aero oOEM", "Other Aero") and deletes visible rows.

 

13. **Save File**: Saves the filtered worksheet as a new file, appending " - Light Version" to the original file name.


