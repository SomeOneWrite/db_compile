from docx import Document
from Model import FileModel, SqlModel
import re
import pdb
from Parsers import ParseWorkers
from Helpers import to_float, to_float_or_zero


class Parse:
    def __init__(self, model):
        self.model = model
        self.last_id = None
        self.collection_id = None
        self.table_aligment_id = {

        }

    def run(self, doc: Document, collection_id: int, id_prefix: str):
        self.collection_id = collection_id
        self.id_prefix = id_prefix
        count = 0
        for table in range(0, len(doc.tables)):
            self.parse_caption(doc.tables[table])
            count += 1

    def get_unit(self, row_str):
        result = re.search(r'(Измеритель:\s*\d{0,10}\s?\b(\w*)\b)', row_str, re.IGNORECASE)
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
                if self.check_podrazdel(cell_text):
                    return row
                if self.check_razdel(cell_text):
                    return row
                if self.check_table_name(cell_text):
                    return row

            if cells[0].text == cells[1].text == cells[2].text == cells[3].text:
                addon_string = cells[0].text
                continue
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
                                result = re.search(r'\d\d-\d\d-\d\d\d-\d\d', txt)
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
                result = re.search(r'\d\d-\d\d-\d\d\d-\d\d', iden)
                if not result:
                    continue
                iden = self.id_prefix + iden
                cost_workers = to_float_or_zero(cells[self.table_aligment_id['4']].text)

                cost_machines = to_float_or_zero(cells[self.table_aligment_id['5']].text)
                cost_drivers = to_float_or_zero(cells[self.table_aligment_id['6']].text)
                cost_materials = to_float_or_zero(cells[self.table_aligment_id['7']].text)
                self.model.insert_unit_position(iden, name, self.unit_name, cost_workers, cost_machines, cost_drivers,
                                                cost_materials, self.last_table_id)
        return row - 2


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
                if self.check_razdel(text_all) and not self.check_podrazdel(text_all):
                    self.last_id = self.model.insert_caption(self.collection_id, None, text)
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
                        continue_n = row - self.parse_unit_position(table, row + 1)
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
                            continue_n = row - self.parse_unit_position(table, row + 1)
                            cells = rows[row].cells
                            break
                        else:
                            self.model.commit()
                            print('text_all = {}'.format(text_all))
                            return
                            pdb.set_trace()
                else:
                    break
