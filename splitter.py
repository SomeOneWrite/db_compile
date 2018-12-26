import sqlite3
import re

def clean(s):
    name_1 = re.sub(r'-', r'', s)
    name_1 = re.sub(r"'", r'', name_1)
    name_1 = re.sub(r':', r'', name_1)
    name_1 = re.sub(r',', r'', name_1)
    name_1 = re.sub(r'\W', r'', name_1)
    name_1 = re.sub(r'\n', r'', name_1)
    name_1 = re.sub(r'\t', r'', name_1)
    name_1 = re.sub(r'\s\s', r'\s', name_1)
    return name_1
omsk_db_name = 'minstroy.odb'
minstroy_db_name = 'minstroy_old.550.odb'
result_db_name = 'r.sqlite'

conn_omsk = sqlite3.connect(omsk_db_name)
conn_min = sqlite3.connect(minstroy_db_name)
conn_result = sqlite3.connect(result_db_name)

query_new = conn_omsk.cursor()
query_old = conn_min.cursor()
query_result = conn_result.cursor()

dict_machines = dict()

res = query_new.execute("select id, name from materials").fetchall()

for r in res:
    dict_machines[r[0]] = r

res = query_old.execute("select id, name from materials").fetchall()
count = 0

for d in dict_machines:
    for r in res:
        name_1 = clean(dict_machines[d][1])
        name_2 = clean(r[1])
        if name_1 == name_2:
            count += 1
            query_old.execute("update materials set id_new = ? where id = ?", [d, r[0]])
            print(count)
            break

conn_omsk.commit()





