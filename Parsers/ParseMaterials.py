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
        self.last = None
        self.last_kniga_id = None
        self.last_chapter_id = None
        self.last_razdel_id = None
        self.last_gruppa_id = None
        self.last_table_id = None

    def run(self, doc: Document, collection_id: int, id_prefix: str, collection_type: int):
        self.collection_id = collection_id
        self.id_prefix = id_prefix
        self.collection_type = collection_type
        count = 0
        for table in range(0, len(doc.tables)):
            self.parse_caption(doc.tables[table])
            count += 1

    def check_razdel(self, row_str: str):
        result = re.search(r'(\W|^)(раздел\s(.*?))(группа)', row_str, re.IGNORECASE)
        if result:
            return result.group(2)
        return None

    def check_table_name(self, row_str: str):
        result = re.search(r'(\W|^)(Таблица\s(.*?))($)', row_str, re.IGNORECASE | re.DOTALL)
        if result:
            return result.group(2)
        return None

    def check_chapter(self, row_str: str):
        result = re.search(r'(\W|^)(часть\s(.*?))(раздел|группа)', row_str, re.IGNORECASE)
        if result:
            return result.group(2)
        return None

    def check_kniga(self, row_str: str):
        result = re.search(r'(\W|^)(книга\s(.*?))(часть|раздел)', row_str, re.IGNORECASE)
        if result:
            return result.group(2)

    def check_gruppa(self, row_str: str):
        result = re.search(r'((Группа(.*?))$)', row_str, re.IGNORECASE)
        if result:
            return result.group(2)
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
        row = 0
        for row in range(row_i, len(rows)):
            if continue_n > 0:
                continue_n -= 1
                continue
            cells = rows[row].cells
            for cell in range(0, len(cells)):
                cell_text = rows[row].cells[cell].text
                if self.check_kniga(cell_text):
                    return row - 1
                if self.check_chapter(cell_text):
                    return row - 1
                if self.check_razdel(cell_text):
                    return row - 1
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
                result = re.search(r'((\d\d\W\d\W\d\d\W\d\d\W\d{2,4})|(\d\d-\d\d-\d\d\d-\d\d))', id)
                if not result:
                    continue
                if addon_string:
                    name = addon_string + ' ' + without_lines(cells[1].text)
                else:
                    name = addon_string + '' + without_lines(cells[1].text)

                unit = without_lines(cells[2].text)
                cost = to_float(cells[3].text)
                cost_smeta = to_float(cells[4].text)
                self.model.insert_material(id, name, unit, cost, cost_smeta, self.last_gruppa_id)
                continue
            if len(rows[row].cells) > 6:
                print('sghafolksfjlhlgaf;gdsjhfj               {}'.format(len(rows[row].cells)))
        return row

    def get_all_text(self, rows, row_i, count: int = 1):
        text_all = ''
        last_text = ''
        is_all = False
        for row in range(row_i, len(rows)):
            self.check_table_aligment(rows[row])
            cells = rows[row].cells
            if is_all:

                return text_all, row
            for i in cells:
                if i.text == last_text:
                    continue
                if re.search(r'(Группа(.*?)\n|\t)', i.text):
                    is_all = True
                last_text = i.text
                text_all += i.text
            if count == 1:
                return text_all, row

        return text_all, row

    def all_checks(self, all_text_):
        all_text = re.sub(r'\n', '', all_text_)
        all_text = re.sub(r'\t', '', all_text)
        name = self.check_kniga(all_text)
        if name:
            self.last_kniga_id = self.model.insert_caption(self.collection_id, None, name)
            self.last = self.last_kniga_id

        name = self.check_chapter(all_text)
        if name:
            self.last_chapter_id = self.model.insert_caption(self.collection_id, self.last_kniga_id, name)
            self.last = self.last_chapter_id

        name = self.check_razdel(all_text)
        if name:
            self.last_razdel_id = self.model.insert_caption(self.collection_id, self.last_chapter_id, name)
            self.last_gruppa_id = None
            self.last = self.last_razdel_id



    def check_captions(self, rows, row_i, table):
        for row in range(row_i, len(rows)):
            self.model.commit()
            text_all, c_row = self.get_all_text(rows, row_i, count=5)
            self.all_checks(text_all)
            self.check_table_aligment(rows[row])
            name = self.check_table_name(text_all)
            if name:
                last = self.get_last_id_for_table()
                self.last_table_id = self.model.insert_caption(self.collection_id, last, name)
                next_row = self.parse_unit_position(table, c_row)
                return next_row
            else:
                name = self.check_gruppa(text_all)
                if name:
                    last = self.get_last_id_for_table()
                    self.last_gruppa_id = self.model.insert_caption(self.collection_id, self.last_razdel_id, name)
                    next_row = self.parse_unit_position(table, c_row - 1)
                    return next_row - 1
        return len(rows)

    def parse_caption(self, table):
        self.unit_name = None
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
            continue_n = self.check_captions(rows, row, table) - row
            self.model.commit()

    def get_last_id_for_table(self):
        if not self.last:
            print('ERROR self.last == None')
        return self.last