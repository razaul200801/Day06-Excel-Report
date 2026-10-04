import openpyxl
from openpyxl.styles import Font, Alignment

input_file = 'sales_data.xlsx'  ## File Path

workbook = openpyxl.load_workbook(input_file)  ## Load the workbook
sheet = workbook.active  ## Get the active sheet

print("Excel File Loaded Successfully!")
print(f"Sheet Name: {sheet.title}")

print("\nReading Data from Excel Sheet:")

for row in sheet.iter_rows(values_only = True):
    print(row)


print("\nCalculating Total Sales for Each Customer:")

for row in sheet.iter_rows(min_row = 2, values_only = True):
    customer = row[0]
    product = row[1]
    quantity = row[2]
    unit_price = row[3]
    total_sales = quantity * unit_price
    print(f"{customer} - {product}: ${total_sales:.2f}")


output_file = 'sales_report.xlsx'  ## Output File Path
report_workbook = openpyxl.Workbook()  ## Create a new workbook
report_sheet = report_workbook.active  ## Get the active sheet
report_sheet.title = "Sales Report"  ## Set the sheet title
report_sheet.append(["Customer", "Product", "Quantity", "Unit Price", "Total Sales"])  ## Add header row

total_sales = []
for row in sheet.iter_rows(min_row = 2, values_only = True):
    customer = row[0]
    product = row[1]
    quantity = row[2]
    unit_price = row[3]
    total_sale = quantity * unit_price
    #print(f"{customer} - {product}: ${total_sale:.2f}")

    next_row = report_sheet.max_row + 1  ## Get the next row index for appending data
    
    total_sales.append(total_sale)  ## Store total sales for later use
    report_sheet.append(
        [customer, 
         product, 
         quantity, 
         unit_price, 
         f"=C{next_row}*D{next_row}"]
        )  ## Append data to report sheet

#Format the header row
for cell in report_sheet[1]:
    cell.font = Font(bold = True)
    cell.alignment = Alignment(horizontal = 'center')


#Set column widths for better readability
report_sheet.column_dimensions['A'].width = 15
report_sheet.column_dimensions['B'].width = 15
report_sheet.column_dimensions['C'].width = 12
report_sheet.column_dimensions['D'].width = 15
report_sheet.column_dimensions['E'].width = 15

#Summary statistics
total_revenue = sum(total_sales)  ## Calculate total revenue
average_sales = total_revenue / len(total_sales)  ## Calculate average sales
highest_sale = max(total_sales)  ## Find the highest sale
lowest_sale = min(total_sales)  ## Find the lowest sale

summary_row = report_sheet.max_row + 2  ## Row index for summary statistics

report_sheet.cell(summary_row, 1, "Sales Summary")
report_sheet.cell(summary_row + 1, 1, "Total Revenue")
report_sheet.cell(summary_row + 1, 2, total_revenue)

report_sheet.cell(summary_row + 2, 1, "Average Sales") 
report_sheet.cell(summary_row + 2, 2, average_sales)

report_sheet.cell(summary_row + 3, 1, "Highest Sale")
report_sheet.cell(summary_row + 3, 2, highest_sale)

report_sheet.cell(summary_row + 4, 1, "Lowest Sale")
report_sheet.cell(summary_row + 4, 2, lowest_sale)

#Formating Summary Statistics
report_sheet.cell(summary_row, 1).font = Font(bold = True)

for row_number in range(summary_row + 1, summary_row + 5):
    report_sheet.cell(row_number, 1).font = Font(bold = True)
    report_sheet.cell(row_number, 2).number_format = '$#,##0.00'  ## Format as currency
    report_sheet.cell(row_number, 2).alignment = Alignment(horizontal = 'left')  ## Align to the left

# Add Auto Filter to the report sheet
report_sheet.auto_filter.ref = "A1:E9"  ## Adjust the range as needed

#Freeze the header row for better navigation
report_sheet.freeze_panes = "A2"  ## Freeze the first row


report_workbook.save(output_file)  ## Save the report workbook
print(f"\nSales report generated and saved as '{output_file}'")


