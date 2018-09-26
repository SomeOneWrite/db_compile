import os
import sqlite3
from docx import Document
from Parsers import ParseWorkers

query_create_tables = [
    'CREATE TABLE dirs ( id integer primary key autoincrement, parent_id integer references dirs(id) on delete cascade on update cascade, name text )',
    'CREATE TABLE collections ( id integer primary key autoincrement, dir_id integer references dirs(id) on delete cascade on update cascade, type integer not null, name text, techpart text )',
    'CREATE TABLE captions ( id integer primary key autoincrement, collection_id integer not null references collections(id) on delete cascade on update cascade, parent_id integer references captions(id) on delete cascade on update cascade, name text )',
    'CREATE TABLE unit_positions ( id text not null, name text, unit text, cost_workers real, cost_machines real, cost_drivers real, cost_materials real, mass real, caption_id integer not null references captions(id) on delete cascade on update cascade )',
    'CREATE TABLE options ( name text not null, val text, primary key(name) )',
    'CREATE TABLE materials ( id text not null, name text, unit text, price real, price_vacantion real, caption_id integer references captions(id) on delete set null on update cascade, primary key(id) )',
    'CREATE TABLE machines ( id text not null, name text, unit text, price real, price_driver real, caption_id integer references captions(id) on delete set null on update cascade, primary key(id) )',
    'CREATE TABLE transports ( id text not null, name text, unit text, price real, type integer check(type >= 1 and type <= 4), caption_id integer references captions(id) on delete set null on update cascade, primary key(id) )',
    'CREATE TABLE workers ( id text not null, name text, unit text, price real, caption_id integer references captions(id) on delete set null on update cascade, primary key(id) )',
]

class Model:
    def __init__(self, db_name):
        self.db_name = db_name
        self.init_db()

    def init_db(self):
        if os.path.exists(self.db_name):
            os.remove(self.db_name)
        self.db_connection = sqlite3.connect(self.db_name)
        self.db_cursor = self.db_connection.cursor()
        for query_create in query_create_tables:
            try:
                self.db_cursor.execute(query_create)
            except sqlite3.DatabaseError as err:
                print("Error: ", err)
            else:
                self.db_connection.commit()
        print("Database and tables created")


    def insert_dir(self, dir_name, dir_id):
        return self.db_cursor.execute('insert into dirs(parent_id, name) values(?, ?)', (dir_id, dir_name))

    def parse_file(self, filename : str, ):
        doc = Document(filename)
        if filename.endswith("workers.docx"):
            ParseWorkers.ParseWorkers(self.db_cursor).run(doc)
            return
        if filename.endswith("transports.docx"):
            ParseWorkers.ParseWorkers(self.db_cursor).run(doc)
            return
        if filename.endswith("machines.docx"):
            ParseWorkers.ParseWorkers(self.db_cursor).run(doc)
            return
        if filename.endswith("materials.docx"):
            ParseWorkers.ParseWorkers(self.db_cursor).run(doc)
            return

        pass



