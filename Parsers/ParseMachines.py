import re

from docx import Document

from Helpers import to_float, without_whitespace, to_float_or, without_lines


class ParseMachines:
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
            self.model.commit()
            for cell in range(0, len(cells)):
                cells = rows[row].cells
                text = cells[cell].text
                text_all = ' '
                for i in cells:
                    text_all += i.text
                text_all = re.sub(r'\n', '', text_all)
                text_all = re.sub(r'\t', '', text_all)
                if self.check_kniga(text_all):
                    self.last_kniga_id = self.model.insert_caption(self.collection_id, None, text)
                    break
                elif self.check_razdel(text_all):
                    self.last_razdel_id = self.model.insert_caption(self.collection_id, self.last_kniga_id, text)
                    break
                elif self.check_gruppa(text_all):
                    if cells[cell].text:
                        self.last_table_id = self.model.insert_caption(self.collection_id, self.last_razdel_id, cells[cell].text)
                    else:
                        self.last_table_id = self.model.insert_caption(self.collection_id, self.last_razdel_id,
                                                                       cells[cell + 1].text)
                    continue_n = row - self.parse_costs(table, row + 1)

                    cells = rows[row].cells
                    break
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
                if self.check_kniga(cell_text):
                    return row - 1
                if self.check_razdel(cell_text):
                    return row - 1
                if self.check_gruppa(cell_text):
                    return row - 1

            if len(rows[row].cells) == 5:
                if cells[1].text == cells[2].text:
                    addon_string = without_lines(cells[1].text)
                    continue
                id = without_lines(without_whitespace(cells[0].text))
                name = addon_string + ' ' + without_lines(cells[1].text)
                unit = without_lines(without_whitespace(cells[2].text))
                price = to_float(cells[3].text)
                price_driver = to_float(cells[4].text)

                self.model.insert_machine(id, name, unit, price, price_driver, self.last_table_id)

                continue
            if len(rows[row].cells) > 6:
                print('sghafolksfjlhlgaf;gdsjhfj               {}'.format(len(rows[row].cells)))
                pass
        return row


    def check_kniga(self, row_str: str):
        row_str = re.sub(r'\n', '', row_str)
        row_str = re.sub(r'\t', '', row_str)
        result = re.search(r'(Книга(.*?)$)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None


    def check_razdel(self, row_str: str):
        result = re.search(r'(Раздел(.*?)$)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None


    def check_gruppa(self, row_str: str):
        result = re.search(r'((Группа(.*?))$)', row_str, re.IGNORECASE)
        if result:
            return result.group(1)
        return None
