import data_generator
import random

class Header:
    def __init__(self):
        self.layout = self.randomize_header_layout()

    def randomize_header_layout():
        order_of_first_four = [1, 2, 3, 4]
        random.shuffle(order_of_first_four)

        possible_header_layouts = [
            [f"B{order_of_first_four[0]}",
             f"B{order_of_first_four[1]}",
             f"B{order_of_first_four[2]}",
             f"B{order_of_first_four[3]}",
             "C1:K4", "L2:M4", "N2:O4",
             f"A{order_of_first_four[0]}",
             f"A{order_of_first_four[1]}",
             f"A{order_of_first_four[2]}",
             f"A{order_of_first_four[3]}",
             "L1:M1", "N1:O1"
             ],
            [f"B{order_of_first_four[0]}",
             f"B{order_of_first_four[1]}",
             f"B{order_of_first_four[2]}",
             f"B{order_of_first_four[3]}",
             "C1:K4", "N2:O4", "L2:M4",
             f"A{order_of_first_four[0]}",
             f"A{order_of_first_four[1]}",
             f"A{order_of_first_four[2]}",
             f"A{order_of_first_four[3]}",
             "N1:O1", "L1:M1"
             ],
            [f"M{order_of_first_four[0]}",
             f"M{order_of_first_four[1]}",
             f"M{order_of_first_four[2]}",
             f"M{order_of_first_four[3]}",
             "C1:K4", "A2:B4", "N2:O4",
             f"K{order_of_first_four[0]}",
             f"K{order_of_first_four[1]}",
             f"K{order_of_first_four[2]}",
             f"K{order_of_first_four[3]}",
             "A1:B1", "N1:O1"
             ],
            [f"M{order_of_first_four[0]}",
             f"M{order_of_first_four[1]}",
             f"M{order_of_first_four[2]}",
             f"M{order_of_first_four[3]}",
             "C1:K4", "N2:O4", "A2:B4",
             f"L{order_of_first_four[0]}",
             f"L{order_of_first_four[1]}",
             f"L{order_of_first_four[2]}",
             f"L{order_of_first_four[3]}",
             "N1:O1", "A1:B1"
             ],
            [f"K{order_of_first_four[0]}",
             f"K{order_of_first_four[1]}",
             f"K{order_of_first_four[2]}",
             f"K{order_of_first_four[3]}",
             "A1:I4", "L2:M4", "N2:O4",
             f"K{order_of_first_four[0]}",
             f"J{order_of_first_four[1]}",
             f"J{order_of_first_four[2]}",
             f"J{order_of_first_four[3]}",
             "L1:M1", "N1:O1"
             ],
            [f"K{order_of_first_four[0]}",
             f"K{order_of_first_four[1]}",
             f"K{order_of_first_four[2]}",
             f"K{order_of_first_four[3]}",
             "A1:I4", "N2:O4", "L2:M4",
             f"J{order_of_first_four[0]}",
             f"J{order_of_first_four[1]}",
             f"J{order_of_first_four[2]}",
             f"J{order_of_first_four[3]}",
             "N1:O1", "L1:M1"
             ],
        ]

        # specify the columns and rows of the header
        order = random.choice(possible_header_layouts)
        formats = []
        header_data = [
            data_generator.HeaderGenerator.generate_date(),
            data_generator.HeaderGenerator.generate_office_code(),
            data_generator.HeaderGenerator.generate_client_code(),
            data_generator.HeaderGenerator.generate_route(),
            data_generator.HeaderGenerator.generate_client_name(),
            data_generator.HeaderGenerator.generate_department_code(),
            data_generator.HeaderGenerator.generate_department_name(),
            "納品日",
            "事業所コード",
            "得意先コード",
            "ルート",
            "部門コード",
            "部門"
        ]
        for i in range(len(order)):
            if(order[i].find(":") != -1):
                if(order[i].find("1") != -1 and order[i].find("4") != -1):
                    formats.append(generate_format(1))
                elif(order[i].find("2") != -1):
                    formats.append(generate_format(2))
                else:
                    formats.append(generate_format())
            else:
                formats.append(generate_format())

        dict = {"cells": order, "data": header_data, "formats": formats}

        return dict
    
class Row:
    def __init__(self):
        self.layout = self.randomize_row_layout()

    def randomize_row_layout():
        formats = []
        standard_format = generate_format()

        cells = ["A6", "B6", "C6", "D6", "E6", "F6", "G6", "I6", "J6", "K6", "L6", "M6", "N6", "O6"]
        row_data = ["マーク","品名","商品コード","数量","単位","金額","通貨","マーク","品名","商品コード","数量","単位","金額","通貨"]
        formats = [standard_format, standard_format, standard_format, standard_format, standard_format, standard_format, standard_format, standard_format, standard_format, standard_format, standard_format, standard_format, standard_format, standard_format]

        rows_in_column_1 = random.randint(1, 25)
        cells_1 = []
        row_data_1 = []
        for i in range(rows_in_column_1):
            # combined cell
            coefficient = 1
            if random.randint(1, 100) <= coefficient:
                cells_1.append(f"A{7 + i}:G{7 + i}")
                isClientCode = random.choice([True, False])
                if(isClientCode):
                    data = f"得意先コード:  {data_generator.HeaderGenerator.generate_client_code()}"
                else:
                    data = f"部門:  {data_generator.HeaderGenerator.generate_department_code()}"
                row_data_1.append(data)
                formats.append(generate_format(3))
            else:
                cells_1.append(f"A{7 + i}")
                cells_1.append(f"B{7 + i}")
                cells_1.append(f"C{7 + i}")
                cells_1.append(f"D{7 + i}")
                cells_1.append(f"E{7 + i}")
                cells_1.append(f"F{7 + i}")
                cells_1.append(f"G{7 + i}")

                row_data_1.append(data_generator.RowGenerator.generate_random_japanese_word(3))
                row_data_1.append(data_generator.RowGenerator.generate_random_japanese_word(10))
                row_data_1.append(data_generator.RowGenerator.generate_product_code())
                row_data_1.append(data_generator.RowGenerator.generate_amount())
                row_data_1.append(data_generator.RowGenerator.generate_amount_unit())
                row_data_1.append(data_generator.RowGenerator.generate_price())
                row_data_1.append(data_generator.RowGenerator.generate_currency())

                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)

        cells.extend(cells_1)
        row_data.extend(row_data_1)

        rows_in_column_2 = random.randint(1, 25)
        cells_2 = []
        row_data_2 = []
        for i in range(rows_in_column_2):
            # combined cell
            coefficient = 1
            if random.randint(1, 100) <= coefficient:
                cells_2.append(f"H{7 + i}:N{7 + i}")
                isClientCode = random.choice([True, False])
                if(isClientCode):
                    data = f"得意先コード:  {data_generator.HeaderGenerator.generate_client_code()}"
                else:
                    data = f"部門:  {data_generator.HeaderGenerator.generate_department_code()}"
                row_data_2.append(data)
                formats.append(generate_format(3))
            else:
                cells_2.append(f"I{7 + i}")
                cells_2.append(f"J{7 + i}")
                cells_2.append(f"K{7 + i}")
                cells_2.append(f"L{7 + i}")
                cells_2.append(f"M{7 + i}")
                cells_2.append(f"N{7 + i}")
                cells_2.append(f"O{7 + i}")

                row_data_2.append(data_generator.RowGenerator.generate_random_japanese_word(3))
                row_data_2.append(data_generator.RowGenerator.generate_random_japanese_word(10))
                row_data_2.append(data_generator.RowGenerator.generate_product_code())
                row_data_2.append(data_generator.RowGenerator.generate_amount())
                row_data_2.append(data_generator.RowGenerator.generate_amount_unit())
                row_data_2.append(data_generator.RowGenerator.generate_price())
                row_data_2.append(data_generator.RowGenerator.generate_currency())

                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)
                formats.append(standard_format)

        cells.extend(cells_2)
        row_data.extend(row_data_2)
        
        dict = {"cells": cells, "data": row_data, "formats": formats}
        return dict

def generate_format(type=0):
    listOfFonts = ['Arial', 'Times New Roman', 'Calibri', 'Verdana', 'Courier New', 'Impact', 'Comic Sans MS', 'Tahoma']
    font = random.choice(listOfFonts)
    if(type == 0):
        font_size = random.randint(6, 10)
    elif(type == 1):
        font_size = random.randint(20, 36)
    elif(type == 2):
        font_size = random.randint(12, 20)
    elif(type == 3):
        font_size = 8
    boldness = random.choice([True, False])
    border = random.randint(1, 5)
    if(type == 0):
        align = random.choice(['left', 'center', 'right'])
    else:
        align = 'center'
    if(type == 0):
        valign = random.choice(['top', 'vcenter', 'bottom'])
    else:
        valign = 'vcenter'
    format = {
        'font': font,
        'font_size': font_size,
        'bold': boldness,
        'border': border,
        'align' : align,
        'valign' : valign
    }
    return format
