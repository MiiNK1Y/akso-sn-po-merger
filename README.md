<div align="center">
    <img src="img/AKSOSPM.png" width="250px" align="center" alt="Logo" />
    <br>
    <br>
    <h1>AKSOSPM</h1>
    <h3>akso-sn-po-manager</h3>
    Bind data from two parallel cells and paste the bound data to the counterpart in a newer sheet.
</div>


## The problem AKSOSPM is made to solve

The system at my previous job did not have a function to bind a PO-number to new products we would receive at our storage facility.
The number was needed for the product-lines in a stock-report (Excel sheet) we would send, once a week, to the customer / the owner of the products, so they could keep tabs with what we had and where to send those specific products, since some locations needed products from some specific pallet.

We solved this in the most round-about way possible;
While in a Excel-sheet; scan every product on a pallet, assign the PO-number to the products on that pallet. And do that once a week, till the end of time... Or till the functionality was implemented into our system... But I wouldn't hold my breath for that to happen.

So, while thinking of a solution, I found that one could interact with an .xlsx sheet with Python through Openpyxl, a tool downloadable with PIP.
After multiple iterations, and working in ConfigParser to the mix, I was left with the program we see here today. Compiled to executables, ready for use with both Windows or Unix / Linux.

TL;DR: We had no implementation to link serial numbers with PO numbers.

## How the program goes about

This explanation is specific to the default config file in the repo. You may, <em>of course</em>, edit the config to fit your need.

1. The program starts by seeking two .xlsx files in the same directory that has dates in their file-names.

    <em>*directory may be edited in the config:<br></em>
    ```ini
    old_sheet_path = ./
    new_sheet_path = ./
    ```
    <em>*date format: dd.mm.yyyy.<br>
    *If there are none or only one file or there is no date that matches the date-format; a custom error is thrown.</em>

2. The program continues by deciding which one of the files are the newest and which is the oldest.

    <em>*This is to decide in what file we should seek the data we want.</em>

3. The program then opens both files, and starts to seek the headings of the oldest file.

    <em>*In this case, we want to find "Serial number" and "PO number".<br>
    *headings to seek may be edited in the config:<br></em>
    ```ini
    serial_column_text = Serial number
    po_column_text = PO kunde
    ```
    <em>*If these aren't found, the program will throw an error.</em>

4. The program will then start to combine the serial number and PO that are found on the same row, in a single list, with a "." (period) splitting the two values.<br>

    <em>*When this process runs, the program will skip any rows that miss a serial number. In this case, that means the product has no serial, and must be an 'ACCESSORIES" -type product, which will be filled in instead of a PO number in the new sheet that will be produced.</em>

5.  When this process is completed, the old file is closed, and we start to work with the newest file.

6.  First off, the system creates a file with some sensitive, internal information for each product, so the columns containing this information needs to be delete before we send it to the customer.

    <em>*See the config file for which columns we will be deleting by index, each column separated by "," (comma) and no white-space between each index-value.<br>
    *These may be edited.</em>
    ```ini
    columns_to_delete = 1,4,5,12,13
    ```

7.  Like with the old sheet, we seek the same heading "Serial number", but since this sheet is straight from the system my job was using, it doesn't have a "PO number" heading, so the program creates that for us.

    <em>*In this case, on column-index "9".<br>
    *Column to insert the PO may be edited in the config.</em>
    ```ini
    final_sheet_column_insert_po = 9
    ```

8.  Then, we start to seek each serial number and then paste the related PO number that was split by a "." (period) on the same row as the serial was found, beneath the "PO number" heading.

    <em>*If there is a product that does not have a serial, then instead of a PO number, the program will fill "ACCESSORIES" in the cell.<br>
    *What to fill in may be edited in the config:</em>
    ```ini
    none_to_match_replacement = ACCESSORIES
    ```
    <em>*In this case, if the serial found was not found in the old sheet, that means the product is new, and the related PO will need to be filled in by hand after the program has ran.<br>
    The program will throw a reminder for this, and how many products needs to have the PO filled in.<br>
    The program will then instead fill in "404_not_found" (for fun) where the PO should be.<br>
    *You can edit the text that will be filled in instead in the config:</em>
    ```ini
    no_match_replacement = 404_not_found
    ```

9.  When the program has reached the bottom of the file, the program will then save this (the newest file) as a new file called "final.xlsx".

    <em>*Which again, can be renamed in the config:</em>
    ```ini
    final_sheet_name = final.xlsx
    ```
    <em>And the file we just worked with (the newest file), will then be closed.</em>

10. You can now take a look at the log the program created, see if there was any errors, or just close the program by pressing [ENTER], since we now are done :)

    <em>*Here, using the demo files found in the "demo_files" directory.<br>
    *If you want to try; remember to copy the two .xlsx files to the same path as the executable.</em>

    ```bash
    ~ $ ./akso-sn-po-merger.py 

    ----------logging started----------
    [+] Found file: test_file_17.05.2024.xlsx
    [+] Found file: test_file_10.05.2024.xlsx
    [+] Formated date: ['2024', '05', '17']
    [+] Formated date: ['2024', '05', '10']
    [+] Newest date found: 17.05.2024
    [+] New sheet assigned to: ./test_file_17.05.2024.xlsx
    [+] Old sheet assigned to: ./test_file_10.05.2024.xlsx
    -----------logging ended-----------

    [!] Serial numbers without PO: 5
    [!] Remember to add the new POs!!
    [!] '404_not_found' is written where the PO should be.

    [+] DONE!

    [x] Press [ENTER] to exit the application.
    ```
