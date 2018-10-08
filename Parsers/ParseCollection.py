from docx import Document
from Model import FileModel, SqlModel
import re
import pdb
from Helpers import to_float, to_float_or_zero, without_whitespace, without_lines
from Parsers.ParseMachines import ParseMachines
from Parsers.ParseMaterials import ParseMaterials
from Parsers.ParseTransports import ParseTransports


iden_reg_pattern = r'((\d\d\W\d\W\d\d\W\d\d\W\d{2,4})|(\d\d-\d\d-\d\d\d-\d\d))|(\d{1,2}-\d{1,2}-\d{1,2})'

class Parse:
    def __init__(self, model):
        self.model = model
        self.last_id = None
        self.collection_id = None
        self.table_aligment_id = {
        }
        self.last_podrazdel_id = None
        self.last_razdel_id = None
        self.last_otdel_id = None
        self.last = None

    def run(self, doc: Document, collection_id: int, id_prefix: str, collection_type : int):
        self.collection_id = collection_id
        self.id_prefix = id_prefix
        self.collection_type = collection_type
        if collection_type == 6:
            ParseTransports(self.model).run(doc, collection_id, id_prefix, collection_type)
            return
        if collection_type == 3:
            ParseMachines(self.model).run(doc, collection_id, id_prefix, collection_type)
            return
        if collection_type == 42:
            ParseMaterials(self.model).run(doc, collection_id, id_prefix, collection_type)
            return
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
        result = re.search(r'(\W|^)(отдел\s(.*?))(раздел)', row_str, re.IGNORECASE)
        if result:
            return result.group(2)

    def check_razdel(self, row_str: str):
        result = re.search(r'(\W|^)(Раздел\s(.*?))(подраздел|таблица)', row_str, re.IGNORECASE)
        if result:
            return result.group(2)
        return None

    def check_podrazdel(self, row_str: str):
        result = re.search(r'(\W|^)(подраздел\s(.*?))(таблица|$)', row_str, re.IGNORECASE)
        if result:
            return result.group(2)
        return None

    def check_table_name(self, row_str: str):
        result = re.search(r'(\W|^)(Таблица\s(.*?))($)', row_str, re.IGNORECASE | re.DOTALL)
        if result:
            return result.group(2)
        return None

    def check_chapter(self, row_str: str):
        result = re.search(r'(\W|^)(часть\s(.*?))(таблица)', row_str, re.IGNORECASE)
        if result:
            return result.group(2)
        return None

    def check_kniga(self, row_str: str):
        row_str = re.sub(r'\n', '', row_str)
        row_str = re.sub(r'\t', '', row_str)
        result = re.search(r'(Книга(.*?)$)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None

    def check_gruppa(self, row_str: str):
        # result = re.search(r'((Группа(.*?))$)', row_str, re.IGNORECASE)
        # if result:
        #     return result.group(2)
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
                cell_text = without_lines(rows[row].cells[cell].text)
                if self.check_podrazdel(cell_text):
                    return row - 1
                if self.check_razdel(cell_text):
                    return row - 1
                if self.check_table_name(cell_text):
                    return row - 1

            if cells[0].text == cells[1].text:
                addon_string = cells[0].text
                continue
            if self.collection_type == 2:
                if len(rows[row].cells) > 7:
                    iden = cells[0].text
                    if to_float(iden):
                        continue
                    if not to_float(cells[2].text):
                        if cells[0].text.strip() == '':
                            txt_all = ''
                            last_str = ''
                            for i in range(0, 4):
                                for j in range(0, 4):
                                    txt = rows[row + i].cells[j].text
                                    result = re.search(iden_reg_pattern, txt)
                                    if result:
                                        continue_n = i
                                        break
                                    if txt != last_str:
                                        txt_all += txt
                                    last_str = txt
                                else:
                                    continue
                                break
                            txt_all = re.sub(r'\n', '', txt_all)
                            txt_all = re.sub(r'\t', '', txt_all)
                            addon_string = txt_all
                            continue
                    name = addon_string + cells[1].text
                    result = re.search(iden_reg_pattern, iden)
                    if not result:
                        continue
                    iden = self.id_prefix + iden
                    cost_workers = to_float_or_zero(cells[self.table_aligment_id['4']].text)

                    cost_machines = to_float_or_zero(cells[self.table_aligment_id['5']].text)
                    cost_drivers = to_float_or_zero(cells[self.table_aligment_id['6']].text)
                    cost_materials = to_float_or_zero(cells[self.table_aligment_id['7']].text)
                    self.model.insert_unit_position(iden, name, self.unit_name, cost_workers, cost_machines, cost_drivers,
                                                    cost_materials, self.last_table_id)
            elif self.collection_type == 4:
                if len(rows[row].cells) > 3:
                    iden = cells[0].text
                    if to_float(iden):
                        continue
                    if not to_float(cells[2].text):
                        if cells[0].text.strip() == '':
                            txt_all = ''
                            last_str = ''
                            for i in range(0, 4):
                                for j in range(0, 4):
                                    txt = rows[row + i].cells[j].text
                                    result = re.search(r'((\d\d\W\d\W\d\d\W\d\d\W\d{2,4})|(\d\d-\d\d-\d\d\d-\d\d))', txt)
                                    if result:
                                        continue_n = i
                                        break
                                    if txt != last_str:
                                        txt_all += txt
                                    last_str = txt
                                else:
                                    continue
                                break
                            txt_all = re.sub(r'\n', '', txt_all)
                            txt_all = re.sub(r'\t', '', txt_all)
                            addon_string = txt_all
                            continue
                    name = addon_string + cells[1].text
                    result = re.search(r'((\d\d\W\d\W\d\d\W\d\d\W\d{2,4})|(\d\d-\d\d-\d\d\d-\d\d))', iden)
                    if not result:
                        continue
                    iden = self.id_prefix + iden
                    cost = to_float_or_zero(cells[self.table_aligment_id['3']].text)
                    cost_smeta = to_float_or_zero(cells[self.table_aligment_id['4']].text)
                    self.model.insert_material(iden, name, self.unit_name, cost, cost_smeta,
                                                        self.last_table_id)
        return row - 2

    def get_all_text(self, rows, row_i, count: int = 1):
        text_all = ''
        last_text = ''
        is_all = False
        for row in range(row_i, len(rows)):
            self.check_table_aligment(rows[row])
            cells = rows[row].cells
            if is_all:
                self.unit_name = re.search(r'Измеритель:\s*(\d{0,10}\s?\b(\w*)\b)', text_all).group(1)
                text_all = re.sub(r'(Измеритель:\s*\d{0,10}\s?\b(\w*)\b)', '', text_all)
                text_all = re.sub(r'\n', '', text_all)
                text_all = re.sub(r'\t', '', text_all)
                return text_all, row
            for i in cells:
                if i.text == last_text:
                    continue
                if re.search(r'(Измеритель:\s*\d{0,10}\s?\b(\w*)\b)', i.text):
                    is_all = True
                last_text = i.text
                text_all += i.text
            text_all = re.sub(r'\n', '', text_all)
            text_all = re.sub(r'\t', '', text_all)
            if count == 1:
                text_all = re.sub(r'\n', '', text_all)
                text_all = re.sub(r'\t', '', text_all)
                return text_all, row
        text_all = re.sub(r'\n', '', text_all)
        text_all = re.sub(r'\t', '', text_all)
        return text_all, row

    def all_checks(self, all_text):
        isChecked = False
        name = self.check_otdel(all_text)
        if name:
            self.last_otdel_id = self.model.insert_caption(self.collection_id, None, name)
            self.last_razdel_id = None
            self.last = self.last_otdel_id
            isChecked = True
        name = self.check_chapter(all_text)
        if name:
            self.last_chast_id = self.model.insert_caption(self.collection_id, None, name)
            self.last = self.last_chast_id
            isChecked = True
        name = self.check_gruppa(all_text)
        if name:
            self.last_gruppa_id = self.model.insert_caption(self.collection_id, None, name)
            self.last = self.last_gruppa_id
            isChecked = True
        name = self.check_razdel(all_text)
        if name:
            if self.last_otdel_id:
                self.last_razdel_id = self.model.insert_caption(self.collection_id, self.last_otdel_id, name)

            else:
                self.last_razdel_id = self.model.insert_caption(self.collection_id, None, name)
            self.last_podrazdel_id = None
            self.last = self.last_razdel_id
            isChecked = True
        name = self.check_podrazdel(all_text)
        if name:
            self.last_podrazdel_id = self.model.insert_caption(self.collection_id, self.last_razdel_id, name)
            self.last = self.last_razdel_id = self.last_podrazdel_id
            isChecked = True
        return isChecked

    def check_captions(self, rows, row_i, table):
        continue_n = 0
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

    def get_last_id_for_table(self):
        if not self.last:
            print('ERROR self.last == None')
        return self.last