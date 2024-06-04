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
The number was needed for the product-lines in a stock-report (Excel sheet) we would send, once a week, to the customer / the owner of the products, so they could keep tabs at what we had and where to send those specific products, since some locations needed products from some specific pallet.

We solved this in the most round-about way possible;
While in a Excel-sheet; scan every product on a pallet, assign the PO-number to the products on that pallet. And do that once a week, till the end of time.
So, while thinking of a solution, I found that one could interact with an .xlsx sheet with Python through Openpyxl, a tool downloadable with PIP.
After multiple iterations, and working in ConfigParser to the mix, I was left with the program we see here today. Compiled to executables, ready for use with both Windows or Unix / Linux.

TL;DR: We had no implementation to link serial numbers with PO numbers.

## How the program goes about

This explanation is specific to the default config file here in the repo, you may, <em>of course</em>, edit the config to fit your need.

1. The program starts by seeking two .xlsx files in the same directory (edit the config to change the working directory) that has dates in their file-names (date format: dd.mm.yyyy).

2. If there are none, only one file, or there is no date that matches the date-format; a custom error is thrown.

3. Else, the program continues by deciding which one of the files are the newest and which is the oldest. (This is to decide from what file we should seek the relevant information).

4. The program then opens both files, and starts to seek the headings of the oldest file.

5. In this case, we want to find "Serial number" and "PO number" (see the config file to change these).

6. If these aren't found, the program will throw an error.

7. Else, the program will then start to combine the serial number and PO that are found on the same row, in a list, with a "." splitting the two values. (Instead of using a dictionary or two lists, this way the data will always share the same index and will never be lost to some sort of bug. And this also allowed me to use Pythons awesome list-comprehension method later in the process, which I was familiar with.)

8. When this process runs, the program will skip any rows that miss a serial number (in this case, that means the product has no serial, and must be an 'ACCESSORIES" -type product, which will be filled in instead of a PO number in the new sheet that will be produced.)

9. When this process is completed, the old file is closed, and we start to work with the newest file.

10. First off, the system creates a file with some sensitive, internal information pr product, so the columns containing this information needs to be delete before we send it to the customer (see the config file for which columns we will be deleting by index, each column separated by "," (comma) and no white-space between each index-value.).

11. Like with the old sheet, we seek the same heading "Serial number", but since this sheet is straight from the system my job was using, it doesn't have a "PO number" heading, so the program creates that for us (in this, on column-index "9", which can be edited in the config.)

12. Then, we start to seek each serial number and then paste the related PO number that was split by a "." (dot) on the same row as the serial was found, beneath the "PO number" heading.

13. If there is a product that does not have a serial, then instead of a PO number, the program will fill "ACCESSORIES" in the cell.

14. And if the serial found was not found in the old sheet, that means the product is new, and the related PO will need to be filled in by hand after the program has ran (the program will throw a reminder for this, and how many products needs to have the PO filled inn.). The program will then instead fill in "404_not_found" (for fun) where the PO should be (you can edit the text that will be filled in instead in the config.)

15. When the program has reached the bottom of the file, the program will then save this (the newest file) as a new file called "final.xlsx" (which again, can be renamed in the config), and the file we just worked with, will then be closed.

16. You can now take a look at the log the program created, see if there was any errors, or just close the program by pressing [ENTER], since we now are done :)
