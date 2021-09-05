# TPR_Reports
A program for processing pricing reports to determine which products with a temporary price reduction(tpr) need to be extended.  Reports are done for every Saturday as it is the end of the companies sale week.

This program was developed for usage at a local grocery store and not intended for widespread usage.

# Purpose

To speed up the process of sorting/filtering the information necessary for us to compile a report of items to review that have expiring pricing.

# Process:

The program gathers information from a spreedsheet that is always titled BRData_Prices.xlsx.  It filters and sorts that data to create a new workbook.

The data it collects is the UPC product code, store department, item description, regular price, and temporarily reduced price(tpr).

The department numbers seperate the UPCs into different sheets.

The UPC's are sorted numerically so that like items can be grouped together for easier to read reports.

After the information is stored in the sheets they are automatically printed and the program closes.