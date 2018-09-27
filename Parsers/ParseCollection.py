from docx import Document
from Model import Model
import re
import pdb
from Parsers import ParseWorkers
from Helpers import to_float, to_float_or_zero


class Parse:
    def __init__(self, model: Model):
        self.model = model
        self.last_id = None
        self.collection_id = None

    def run(self, doc: Document, collection_id: int):
        self.collection_id = collection_id
        count = 0
        for table in range(0, len(doc.tables)):
            self.parse_caption(doc.tables[table])
            count += 1
        pass

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

    def parse_unit_position(self, table, row_i):
        rows = table.rows
        for row in range(row_i, len(rows)):
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
                self.addon_string = cells[0].text
            else:
                self.addon_string = ''
            if len(rows[row].cells) > 7:
                iden = cells[0].text
                if to_float(iden):
                    continue
                name = self.addon_string + cells[1].text
                cost_workers = to_float_or_zero(cells[3].text)
                if not to_float(cells[2].text):
                    continue
                cost_machines = to_float_or_zero(cells[4].text)
                cost_drivers = to_float_or_zero(cells[5].text)
                cost_materials = to_float_or_zero(cells[6].text)
                self.model.insert_unit_position(iden, name, self.unit_name, cost_workers, cost_machines, cost_drivers,
                                                cost_materials, self.last_table_id)

    return row


def parse_caption(self, table):
    rows = table.rows
    for row in range(0, len(rows)):
        cells = rows[row].cells
        for cell in range(0, len(cells)):
            text = cells[cell].text
            text_all = ' '
            for i in cells:
                text_all += i.text
            text_all = re.sub(r'\n', '', text_all)
            text_all = re.sub(r'\t', '', text_all)
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
                self.last_table_id = self.model.insert_caption(self.collection_id, self.last_id, table_name)
                self.unit_name = self.get_unit(text_all)
                if self.unit_name:
                    row = self.parse_unit_position(table, row + 1)
                    cells = rows[row].cells
                    break
                else:
                    self.model.db_connection.commit()
                    print('text_all = {}'.format(text_all))
                    exit(0)
                    pdb.set_trace()
