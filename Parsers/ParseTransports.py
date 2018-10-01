import re

from docx import Document

from Helpers import to_float, without_whitespace, to_float_or


class ParseTransports:
    def __init__(self, model):
        self.model = model
        self.last_id = None
        pass

    def run(self, doc: Document, collection_id: int, id_prefix: str, collection_type: int):
        self.doc = doc
        self.collection_id = collection_id
        self.id_prefix = id_prefix
        self.collection_type = collection_type
        for table in doc.tables:
            self.parse_caption(table)

    def parse_caption(self, table):
        rows = table.rows
        continue_n = 0
        for row in range(0, len(rows)):
            if continue_n > 0:
                continue_n -= 1
                continue
            try:
                cells = rows[row].cells
            except Exception as e:
                continue
            for cell in range(0, len(cells)):
                cells = rows[row].cells
                text = cells[cell].text
                text_all = ' '
                for i in cells:
                    text_all += i.text
                text_all = re.sub(r'\n', '', text_all)
                text_all = re.sub(r'\t', '', text_all)
                if self.check_otdel(text_all):
                    self.last_id = self.model.insert_caption(self.collection_id, None, text)
                    break
                elif self.check_razdel(text_all) and not self.check_podrazdel(text_all):
                    self.last_id = self.model.insert_caption(self.collection_id, self.last_id, text)
                    break
                elif self.check_podrazdel(text_all):
                    self.last_id = self.model.insert_caption(self.collection_id, self.last_id, text)
                    break
                elif self.check_table_name(text_all):
                    table_name = re.sub(r'(Измеритель:\s*\d{0,10}\s?\b(\w*)\b)', '', text, re.IGNORECASE)
                    table_name = re.sub(r'\t', '', table_name)
                    table_name = re.sub(r'\n', '', table_name)
                    self.unit_name = self.get_unit(text_all)
                    if self.unit_name:
                        self.last_table_id = self.model.insert_caption(self.collection_id, self.last_id, table_name)
                        continue_n = row - self.parse_costs(table, row + 1)
                        cells = rows[row].cells
                        break
                    else:
                        table_name = ''
                        for i in range(0, 2):
                            for j in range(0, 3):
                                table_name += rows[row + i].cells[j].text
                        table_name = re.sub(r'\n', '', table_name)
                        table_name = re.sub(r'\t', '', table_name)
                        self.unit_name = self.get_unit(table_name)
                        if self.unit_name:
                            table_name = re.sub(r'(\s*Измеритель:\s*\d{0,10}\s?\b(\w*)\b\s*)', '', table_name,
                                                re.IGNORECASE)
                            self.last_table_id = self.model.insert_caption(self.collection_id, self.last_id, table_name)
                            continue_n = row - self.parse_costs(table, row + 1)
                            cells = rows[row].cells
                            break
                        else:
                            self.model.commit()
                            print('text_all = {}'.format(text_all))
                            return
                else:
                    break

    def parse_costs(self, table, row_i):
        addon_string = ''
        rows = table.rows
        continue_n = 0
        codes = None
        for row in range(row_i, len(rows)):
            if continue_n > 0:
                continue_n -= 1
                continue
            cells = rows[row].cells
            for cell in range(0, len(cells)):
                cell_text = rows[row].cells[cell].text
                if self.check_podrazdel(cell_text):
                    return row
                if self.check_razdel(cell_text):
                    return row
                if self.check_table_name(cell_text):
                    return row

            if len(rows[row].cells) == 4:
                if cells[0].text == cells[1].text == cells[2].text == cells[3].text:
                    addon_string = cells[1].text
                    continue
                if not codes:
                    codes = [cells[0].text, cells[1].text]
                    continue
                id_1 = without_whitespace(codes[0]) + without_whitespace(cells[0].text)
                id_2 = without_whitespace(codes[1]) + without_whitespace(cells[0].text)
                name_1 = addon_string + ' ' + cells[1].text + ', погрузка'
                name_2 = addon_string + ' ' + cells[1].text + ', разгрузка'
                price_1 = to_float(cells[2].text)
                if not price_1:
                    print('error {}', id_1)
                    raise Exception('Error {}'.format(id_1))

                price_2 = to_float_or(cells[3].text, price_1)
                self.model.insert_transport(id_1, name_1, self.unit_name, price_1, None, self.last_table_id)
                self.model.insert_transport(id_2, name_2, self.unit_name, price_2, None, self.last_table_id)
                continue
            if len(rows[row].cells) == 6:
                if cells[0].text == cells[1].text == cells[2].text == cells[3].text:
                    addon_string = (cells[1].text)
                    continue
                if not codes:
                    codes = [cells[0].text, cells[1].text, cells[2].text, cells[3].text]
                    continue
                id_1 = without_whitespace(codes[0]) + without_whitespace(cells[0].text)
                id_2 = without_whitespace(codes[1]) + without_whitespace(cells[0].text)
                id_3 = without_whitespace(codes[2]) + without_whitespace(cells[0].text)
                id_4 = without_whitespace(codes[3]) + without_whitespace(cells[0].text)
                name_1 = addon_string + ' ' + (cells[1].text) + ', I класс груза'
                name_2 = addon_string + ' ' + (cells[2].text) + ', II класс груза'
                name_3 = addon_string + ' ' + (cells[1].text) + ', III класс груза'
                name_4 = addon_string + ' ' + (cells[1].text) + ', IV класс груза'
                price_1 = to_float(cells[2].text)
                if not price_1:
                    print('error {}', id_1)
                    raise Exception('Error {}'.format(id_1))

                price_2 = to_float_or(cells[3].text, price_1)
                price_3 = to_float_or(cells[4].text, price_2)
                price_4 = to_float_or(cells[5].text, price_3)
                self.model.insert_transport(id_1, name_1, self.unit_name, price_1, 1, self.last_table_id)
                self.model.insert_transport(id_2, name_2, self.unit_name, price_2, 2, self.last_table_id)
                self.model.insert_transport(id_3, name_3, self.unit_name, price_3, 3, self.last_table_id)
                self.model.insert_transport(id_4, name_4, self.unit_name, price_4, 4, self.last_table_id)
                continue
            if len(rows[row].cells) > 6:
                print('sghafolksfjlhlgaf;gdsjhfj               {}'.format(len(rows[row].cells)))
                pass
        return 0
        pass

    def check_otdel(self, row_str: str):
        row_str = re.sub(r'\n', '', row_str)
        row_str = re.sub(r'\t', '', row_str)
        result = re.search(r'(Отдел\s?\d{1,2}(.*?))$', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

    def check_razdel(self, row_str: str):
        result = re.search(r'(раздел\s?\d?)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

    def check_podrazdel(self, row_str: str):
        result = re.search(r'(Подраздел\s?\d?)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

    def check_razdel(self, row_str: str):
        result = re.search(r'(раздел\s?\d?)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

    def check_table_name(self, row_str: str):
        result = re.search(r'(Таблица\s?\d?)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

    def get_unit(self, row_str):
        result = re.search(r'(Измеритель:\s*\d{0,10}\s?\b(\w*)\b)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None
