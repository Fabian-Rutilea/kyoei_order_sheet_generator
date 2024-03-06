import random
import calendar
import string

class HeaderGenerator:

    def __init__(self):
        self.date = self.generate_date()
        self.office_code = self.generate_office_code()
        self.client_code = self.generate_client_code()
        self.route = self.generate_route()
        self.client_name = self.generate_client_name()
        self.department_code = self.generate_department_code()
        self.department_name = self.generate_department_name()

    # generate 納品日
    def generate_date():
        year = random.randint(2020, 2100)
        month = random.randint(1, 12)
        max_day = calendar.monthrange(year, month)[1]
        day = random.randint(1, max_day)
        date = f"{month:02d}月{day:02d}日"
        return date

    # generate 事業所コード
    def generate_office_code():
        office_code = ''
        length = random.randint(1, 3)
        for i in range(length):
            office_code += str(random.randint(0, 9))
        return office_code

    # generate 得意先コード
    def generate_client_code():
        client_code = "0"
        length = random.randint(4, 5)
        for i in range(length):
            client_code += str(random.randint(0, 9))
        return client_code

    # generate ルート
    def generate_route():
        route_letter = random.choice('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
        route_number = str(random.randint(0, 9))
        route = f"{route_letter}-{route_number}"
        return route

    # generate 得意先名称
    def generate_client_name():
        company_prefix = random.choice(['株式会社', '有限会社', ''])
        company_name = random.choice(['Toyota', '共栄エンジニアリング', 'Rutilea', 'Apple', '杉本'])
        company_suffix = random.choice(['Ltd.', 'Co.', 'Corp.', 'Inc.', ''])
        client_name = f"{company_prefix}{company_name}{company_suffix}"
        return client_name

    # generate 部門コード
    def generate_department_code():
        department_code = ""
        length = random.randint(0, 4)
        for i in range(length):
            department_code += str(random.randint(0, 9))
        return department_code

    # generate 部門
    def generate_department_name():
        department_name = random.choice(['本部', '関西支店', '仙台支店', '大阪支店', ''])
        return department_name
                
class RowGenerator:

    def __init__(self):
        self.mark = self.generate_random_japanese_word(3)
        self.product_name = self.generate_random_japanese_word(10)
        self.product_code = self.generate_product_code()
        self.amount = self.generate_amount()
        self.amount_unit = self.generate_amount_unit()
        self.price = self.generate_price()
        self.currency = self.generate_currency()

    def generate_product_code():
        product_code = ''
        for i in range(6):
            product_code += str(random.randint(0, 9))
        return product_code
    
    def generate_amount():
        amount = random.choice(['', str(random.randint(0, 99))])
        return amount
    
    def generate_amount_unit():
        amount_unit = random.choice(['CS', 'BL', '本'])
        return amount_unit
    
    def generate_price():
        price = str(random.randint(0, 99999))
        return price

    def generate_currency():
        currency = random.choice(['$', '€', '円'])
        return currency

    def generate_random_english_word(max_length):
        word_length = random.randint(1, max_length)
        return ''.join(random.choice(string.ascii_lowercase) for _ in range(word_length))

    def generate_random_japanese_word(max_length):
        hiragana_characters = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん'

        word_length = random.randint(1, max_length)
        return ''.join(random.choice(hiragana_characters) for _ in range(word_length))