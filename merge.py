import sqlite3
import re
conn = sqlite3.connect('res.sqlite')
query = conn.cursor()
captions = query.execute("select * from captions").fetchall()

last_razdel_id = None
last_podrazdel_id = None

for caption in captions:
    name = caption[3]
    parent_id = caption[2]
    cur_id = caption[0]
    res = re.search(r'Подраздел', name)
    if res:
        last_podrazdel_id = cur_id
        if last_razdel_id:
            query.execute("update captions set parent_id = ? where id = ?", (last_razdel_id, cur_id))
        continue
    res = re.search(r'Раздел', name)
    if res:
        last_razdel_id = cur_id
        continue

        #print("name = {}, id = {}, replaced name = {}", collection[3], collection[1], collection)
        #query.execute("update collections set name = ? where id = ?", (caption[3].replace(".docx", ""), collection[0]))

conn.commit()

# for l in new_list:
#     query.execute('update captions set name = ? where id = ?', (l[3], l[0]))
#
# conn.commit()
#
#
# file_name = '1.docx'
#
# doc = Document(file_name)
#
# tables = doc.tables
#
# for table in tables:
#     rows = table.rows
#     for row in rows:
#         cells = row.cells
#         for cell in cells:
#             tc = cell._tc
#             tcPr = tc.get_or_add_tcPr()
#             tcBorders = tcPr.first_child_found_in("w:tcBorders")
#             if tcBorders is None:
#                 continue
#             tcBorder.values()
#             for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
#                 tag = 'w:{}'.format(edge)
#                 element = tcBorders.find(qn(tag))
#                 if element:
#                     pass

# conn = sqlite3.connect('result.sqlite')
#
# query_unit = conn_unit_pos.cursor()
# query = conn.cursor()
#
# all_collections = query_unit.execute("select * from collections").fetchall()
# last_collection_id = query.execute("select * from sqlite_sequence where name = 'collections'").fetchone()[1]
# all_captions = query_unit.execute("select * from captions").fetchall()
# last_caption_id = query.execute("select * from sqlite_sequence where name = 'captions'").fetchone()[1]
# all_unit_positions = query_unit.execute("select * from unit_positions").fetchall()
# for collection in all_collections:
#     query.execute('insert into collections values(?, ?, ?, ?, ?)', (collection[0] + last_collection_id, collection[1], collection[2], collection[3], collection[4]))
#
# for caption in all_captions:
#     captions_s = list()
#     captions_s.append(caption[0] + last_caption_id)
#     captions_s.append(caption[1] + last_collection_id)
#     if caption[2]:
#         captions_s.append(caption[2] + last_caption_id)
#     else:
#         captions_s.append(caption[2])
#     captions_s.append(caption[3])
#     query.execute('insert into captions values (?, ?, ?, ?)', captions_s)
#
# for unit in all_unit_positions:
#     query.execute('insert into unit_positions values (?, ?, ?, ?, ?, ?, ?, ?, ?)', (unit[0], unit[1], unit[2], unit[3], unit[4], unit[5], unit[6], unit[7], unit[8] + last_caption_id))
#
# conn.commit()
