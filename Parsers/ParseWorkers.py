
class ParseWorkers:
    def __init__(self, query):
        self.query = query

    def check_worker_key(self, cells):
        for cell in cells:
            if cell.text.lower().find('1') != -1:
                return True
        return False

    def parse_prof(self, table):
        for row in range(1, len(table.rows)):
            cells = table.rows[row].cells
            for tmp_cell in range(0, len(cells), 2):
                cell = cells[tmp_cell]
                price = cells[tmp_cell + 1].text
                price = ''.join(price.split())
                price = price.replace(',', '.')
                price = float(price)
                text = cell.text
                text = ''.join(text.strip())
                db_result = self.query.execute("update workers set price = ? where name = ?", [price, text])

    def run(self, doc):
        for table in doc.tables:
            cells = table.rows[0].cells
            if cells[0].text.find('3') != -1:
                self.parse_prof(table.table)
                return
            prefix = cells[0].text
            for row in range(1, len(table.rows)):
                cells = table.rows[row].cells
                for tmp_cell in range(0, len(cells) - 1, 2):
                    try:
                        cell = cells[tmp_cell]
                        text = cell.text
                        text = text.replace('.', '-')
                        text = prefix + '-' + text
                        text = ''.join(text.split())
                        price = ''.join(cells[tmp_cell + 1].text.split())
                        price = price.replace(',', '.')
                        price = float(price)
                        db_result = self.query.execute("select * from workers where workers.id = ?", [text]).fetchall()
                        if db_result:
                            db_result = self.query.execute('update workers set price = ? where workers.id = ?',
                                                           [price, text]).fetchone()
                            if not db_result:
                                print('Cant write with id {} price {}'.format(text, price))
                            else:
                                print('id = {} price = {}'.format(text, price))
                    except:
                        continue
        return
