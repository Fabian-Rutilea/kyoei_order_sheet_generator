import pandas as pd
import matplotlib.pyplot as plt
import os
import sheet_layout
import xlsxwriter

# Function to generate a randomized Excel sheet
def generate_excel_sheet(output_name):
    workbook = xlsxwriter.Workbook(output_name)
    worksheet = workbook.add_worksheet()

    header = sheet_layout.Header.randomize_header_layout()
    header_cells =  header["cells"]
    header_data = header["data"]
    header_formats = header["formats"]

    for i in range(len(header_cells)):
        cell_format = workbook.add_format(header_formats[i])
        if(header_cells[i].find(":") != -1):
            worksheet.merge_range(header_cells[i], header_data[i], cell_format)
        else:
             worksheet.write(header_cells[i], header_data[i], cell_format)

    rows = sheet_layout.Row.randomize_row_layout()
    row_cells =  rows["cells"]
    row_data = rows["data"]
    row_formats = rows["formats"]

    for i in range(len(row_cells)):
        cell_format = workbook.add_format(row_formats[i])
        if(row_cells[i].find(":") != -1):
            worksheet.merge_range(row_cells[i], row_data[i], cell_format)
        else:
            worksheet.write(row_cells[i], row_data[i], cell_format)

    set_column_widths(worksheet)
    set_row_heights(worksheet)

    workbook.close()

def set_column_widths(worksheet):
    worksheet.set_column(0, 0, 4)
    worksheet.set_column(1, 1, 12)
    worksheet.set_column(2, 2, 6)
    worksheet.set_column(3, 3, 3)
    worksheet.set_column(4, 4, 3)
    worksheet.set_column(5, 5, 3)
    worksheet.set_column(6, 6, 3)
    worksheet.set_column(7, 7, 7)
    worksheet.set_column(8, 8, 12)
    worksheet.set_column(9, 9, 6)
    worksheet.set_column(10, 10, 3)
    worksheet.set_column(11, 11, 3)
    worksheet.set_column(12, 12, 3)
    worksheet.set_column(13, 13, 12)
    worksheet.set_column(14, 14, 12)

def set_row_heights(worksheet):
    worksheet.set_row(5, 30)
    worksheet.set_row(6, 30)
    worksheet.set_row(7, 30)
    worksheet.set_row(8, 30)
    worksheet.set_row(9, 30)
    worksheet.set_row(10, 30)
    worksheet.set_row(11, 30)
    worksheet.set_row(12, 30)
    worksheet.set_row(13, 30)
    worksheet.set_row(14, 30)
    worksheet.set_row(15, 30)
    worksheet.set_row(16, 30)
    worksheet.set_row(17, 30)
    worksheet.set_row(18, 30)
    worksheet.set_row(19, 30)
    worksheet.set_row(20, 30)
    worksheet.set_row(21, 30)
    worksheet.set_row(22, 30)
    worksheet.set_row(23, 30)
    worksheet.set_row(24, 30)
    worksheet.set_row(25, 30)
    worksheet.set_row(26, 30)

# Function to convert Excel sheet to JPEG
def excel_to_jpeg(excel_path, jpeg_path):
    df = pd.read_excel(excel_path)

    fig, ax = plt.subplots()
    ax.axis('tight')
    ax.axis('off')
    ax.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')

    plt.savefig(jpeg_path, bbox_inches='tight', pad_inches=0.1)
    plt.close()

# Create CSV and JPEG folders if they don't exist
csv_folder = 'CSV'
jpeg_folder = 'JPEG'
os.makedirs(csv_folder, exist_ok=True)
os.makedirs(jpeg_folder, exist_ok=True)

# Number of Excel sheets to generate
num_sheets = 10

# Generate randomized Excel sheets and convert them to JPEG
for i in range(num_sheets):
    excel_filename = os.path.join(csv_folder, f'sheet_{i+1}.xlsx')
    jpeg_filename = os.path.join(jpeg_folder, f'sheet_{i+1}.jpeg')

    generate_excel_sheet(excel_filename)
    # excel_to_jpeg(excel_filename, jpeg_filename)

print(f'{num_sheets} Excel sheets and JPEG files generated successfully.')
