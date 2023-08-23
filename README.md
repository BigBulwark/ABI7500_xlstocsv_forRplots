# ABI7500 Fast/Standard .xls to .csv Converter for R

**Summary**:

This VBA script facilitates the conversion of exported qPCR data (either multiplex or singleplex) from the ABI7500 Fast/Standard instrument into a specific .csv format. This format is tailored for compatibility with R programming, enabling the creation of various plots based on cycle threshold values (Ct/Cq), such as linear curves, boxplots, and more. The script also ensures that the converted .csv output is saved in a designated file location, retaining the original .xls file name.

**Detailed Workflow**:

1. Worksheets Utilized:
The script uses "Results" as wsInput (the primary data source).
A new worksheet named "R_Output" is created, serving as wsOutput where the reformatted data is stored.

2. File Path and Name Extraction:
The complete file path is sourced from cell B3 in wsInput.
The filename is then extracted from this path.

3. Header Definition:
Headers for the new sheet are predefined and positioned in the first row of wsOutput.

4. Data Iteration and Extraction:
The macro iterates from row 9 to the last data-containing row in column 1 of wsInput.
If both columns 1 and 2 in a row contain data, the corresponding details (including the experiment file name, well, sample name, target name, and Ct) are transferred to wsOutput.

5. CSV File Saving:
wsOutput is saved as a CSV file in the specified location (e.g., "C:\Example"), with the filename derived from expFileName.

6. Workbook Closure:
After exporting the data to a CSV file, the newly created workbook is closed without saving any changes.

