from docx import Document

class Parse:
    def __init__(self, query):
        self.query = query

    def run(self, doc : Document):
        pass

    def insert_caption(self, collection_id : int, parent_id : int, name : str):
        pass

    def insert_collection(self, dir_id : int, type : int, name : str, tech_part : str):
        pass