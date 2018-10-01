from docx import Document
import re
from Helpers import to_float, to_float_or_zero, without_lines, without_whitespace


class ParseMaterials:
    def __init__(self, model):
        self.model = model
        self.last_id = None
        self.collection_id = None
        self.table_aligment_id = {
        }

    def run(self, doc: Document, collection_id: int, id_prefix: str, collection_type: int):
        self.collection_id = collection_id
        self.id_prefix = id_prefix
        self.collection_type = collection_type
        count = 0
        for table in range(0, len(doc.tables)):
            self.parse_caption(doc.tables[table])
            count += 1

    def get_unit(self, row_str):
        result = re.search(r'(Измеритель:\s*\d{0,10}\s?\b(\w*)\b)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

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

    def check_chapter(self, row_str: str):
        result = re.search(r'(Часть\s\d{1,3}\W)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

    def check_kniga(self, row_str: str):
        row_str = re.sub(r'\n', '', row_str)
        row_str = re.sub(r'\t', '', row_str)
        result = re.search(r'(Книга(.*?)$)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

    def check_gruppa(self, row_str: str):
        result = re.search(r'(Группа(.*?)$)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

    def check_table_aligment(self, row):
        for cell in row.cells:
            result = re.sub(r'\d*', '', cell.text).strip()
            if result != '':
                return
        for cell in range(0, len(row.cells)):
            result = row.cells[cell].text.strip()
            self.table_aligment_id[result] = cell

    def parse_unit_position(self, table, row_i):
        addon_string = ''
        rows = table.rows
        continue_n = 0
        for row in range(row_i, len(rows)):
            if continue_n > 0:
                continue_n -= 1
                continue
            cells = rows[row].cells
            for cell in range(0, len(cells)):
                cell_text = rows[row].cells[cell].text
                if self.check_kniga(cell_text):
                    return row
                if self.check_chapter(cell_text):
                    return row
                if self.check_razdel(cell_text):
                    return row
                if self.check_gruppa(cell_text):
                    return row
            if cells[1].text == cells[2].text or cells[1].text == cells[0].text:
                addon_string = without_lines(cells[1].text)
                continue
            if len(rows[row].cells) == 5:
                if cells[1].text == cells[2].text:
                    addon_string = without_lines(cells[1].text)
                    continue
                id = without_lines(without_whitespace(cells[0].text))
                if addon_string:
                    name = addon_string + ' ' + without_lines(cells[1].text)
                else:
                    name = addon_string + '' + without_lines(cells[1].text)
                unit = without_lines(cells[2].text)
                cost = to_float(cells[3].text)
                cost_smeta = to_float(cells[4].text)
                self.model.insert_material(id, name, unit, cost, cost_smeta, self.last_table_id)
                continue
            if len(rows[row].cells) > 6:
                print('sghafolksfjlhlgaf;gdsjhfj               {}'.format(len(rows[row].cells)))
        return row

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
                self.check_table_aligment(rows[row])
                if self.check_kniga(text_all):
                    self.last_id = self.model.insert_caption(self.collection_id, None, text)
                    break
                elif self.check_chapter(text_all):
                    self.last_id = self.model.insert_caption(self.collection_id, self.last_id, text)
                    break
                elif self.check_razdel(text_all):
                    self.last_id = self.model.insert_caption(self.collection_id, self.last_id, text)
                    break
                elif self.check_gruppa(text_all):
                    self.last_table_id = self.model.insert_caption(self.collection_id, self.last_id, text)
                    con = self.parse_unit_position(table, row + 1)
                    continue_n = con - row - 1
                    cells = rows[row].cells
                    break
