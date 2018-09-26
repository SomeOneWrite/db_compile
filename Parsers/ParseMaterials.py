import re

from Helpers import to_float


class ParseMaterials:
    def __init__(self, query):
        self.query = query

    def get_material_price(self, cells):
        result = None
        result_2 = None
        for i in range(1, 6):
            if len(cells) - 1 < i:
                break
            result = cells[i].text
            result = ''.join(result.split())
            result = result.replace(',', '.')
            try:
                result = float(result)
                result_2 = cells[i + 1].text
                result_2 = ''.join(result_2.split())
                result_2 = result_2.replace(',', '.')
                result_2 = float(result_2)
                return result, result_2
            except Exception as e:
                result = None
                continue
            break
        return result, result_2

    def check_material_key(self, cells):
        cell = cells[0]
        result = re.search(r'(\d\d\d-\d\d\d\d)\D', cell.text)
        if result:
            return result.group(1)
        return None

    def check_material_key_text(self, text):
        result = re.search(r'\A(\d\d\d-\d\d\d\d)\s\D', text)
        if result:
            return result.group(1)
        return None

    def run(self, row):
        result = self.check_material_key(row.cells)
        if not result: return
        db_result = self.query.execute("select * from materials where materials.id = ?", [result]).fetchone()
        if db_result:
            price, price_2 = self.get_material_price(row.cells)
            if re.match("^\d+?\.\d+?$", str(price)) is None:
                pass
            if not price:
                print('error update materials with id {}'.format(db_result[0]))
                return
            result = self.query.execute('update materials set price_vacantion = ? where materials.id = ?',
                                        (price, db_result[0]))
            if not result:
                print('error update materials with id {} , not result'.format(db_result[0]))
            return
        return

    def run_text(self, doc, count):
        i = 0  # (\d{0,3}\s{0,1}\d{0,3},\d{1,2})\D{1,2}(\d{0,3}\s{0,1}\d{0,3},\d{1,2})
        pattern = r'(\d{0,3}\s{0,1}\d{0,3},\d{1,2})\D{1,2}(\d{0,3}\s{0,1}\d{0,3},\d{1,2})'
        p_iter = iter(doc.paragraphs)
        ln = len(doc.paragraphs)
        for p in p_iter:
            i += 1
            key = self.check_material_key_text(p.text)
            if key:
                count += 1
                db_result = self.query.execute("select * from materials where materials.id = ?", [key]).fetchone()
                if not db_result:
                    print('these id doesnt exists {} '.format(key))
                else:
                    if db_result[3] != None:
                        continue
                    price = re.search(pattern, p.text)
                    if price:
                        res = self.query.execute(
                            'update materials set price = ?, price_vacantion = ? where materials.id = ?',
                            [to_float(price.group(2)), to_float(price.group(1)), db_result[0]])

                    else:
                        try:
                            p_next = next(p_iter)
                        except:
                            continue
                        key_2 = self.check_material_key_text(p_next.text)
                        if not key_2:
                            price = re.search(
                                pattern,
                                p_next.text)
                            if price:
                                self.query.execute(
                                    'update materials set price = ?, price_vacantion = ? where materials.id = ?',
                                    [to_float(price.group(2)), to_float(price.group(1)), db_result[0]])
                            else:
                                try:
                                    p_next = next(p_iter)
                                except:
                                    continue
                                key_2 = self.check_material_key_text(p_next.text)
                                if not key_2:
                                    price = re.search(
                                        pattern,
                                        p_next.text)
                                    if price:
                                        self.query.execute(
                                            'update materials set price = ?, price_vacantion = ? where materials.id = ?',
                                            [to_float(price.group(2)), to_float(price.group(1)), db_result[0]])
                                    else:
                                        try:
                                            p_next = next(p_iter)
                                        except:
                                            continue
                                        key_2 = self.check_material_key_text(p_next.text)
                                        if not key_2:
                                            price = re.search(
                                                pattern,
                                                p_next.text)
                                            if price:
                                                self.query.execute(
                                                    'update materials set price = ?, price_vacantion = ? where materials.id = ?',
                                                    [to_float(price.group(2)), to_float(price.group(1)), db_result[0]])
                                            else:
                                                try:
                                                    p_next = next(p_iter)
                                                except:
                                                    continue
                                                key_2 = self.check_material_key_text(p_next.text)
                                                if not key_2:
                                                    price = re.search(
                                                        pattern,
                                                        p_next.text)
                                                    if price:
                                                        self.query.execute(
                                                            'update materials set price = ?, price_vacantion = ? where materials.id = ?',
                                                            [to_float(price.group(2)), to_float(price.group(1)),
                                                             db_result[0]])
                                                    else:
                                                        try:
                                                            p_next = next(p_iter)
                                                        except:
                                                            continue
                                                        key_2 = self.check_material_key_text(p_next.text)
                                                        if not key_2:
                                                            price = re.search(
                                                                pattern,
                                                                p_next.text)
                                                            if price:
                                                                self.query.execute(
                                                                    'update materials set price = ?, price_vacantion = ? where materials.id = ?',
                                                                    [to_float(price.group(2)), to_float(price.group(1)),
                                                                     db_result[0]])
                                                            else:
                                                                try:
                                                                    p_next = next(p_iter)
                                                                except:
                                                                    continue
                                                                key_2 = self.check_material_key_text(p_next.text)
                                                                if not key_2:
                                                                    price = re.search(
                                                                        pattern,
                                                                        p_next.text)
                                                                    if price:
                                                                        self.query.execute(
                                                                            'update materials set price = ?, price_vacantion = ? where materials.id = ?',
                                                                            [to_float(price.group(2)),
                                                                             to_float(price.group(1)),
                                                                             db_result[0]])
                                                                    else:
                                                                        print('error updating with id {}'.format(key))
