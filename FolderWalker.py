import os
import Model
from Parsers.ParseCollection import Parse
from docx import Document
import json

class FolderWalker:
    def __init__(self, root_folder):
        self.model = Model.Model("result.sqlite")
        self.parser = Parse(self.model)
        self.root_folder = root_folder
        self.walk(root_folder)
        self.current_collection_type = None
        pass

    def walk(self, folder, dir_id=None):
        for file in os.listdir(folder):
            if (file.startswith('~')): continue
            if file.endswith(".docx"):
                with open(os.path.join(folder, 'config.cfg')) as f:
                    config_file = json.load(f)
                    self.current_collection_type = config_file['type']
                    print('collection type: {}'.format(self.current_collection_type))
                print("File {}".format(file))
                collection_id = self.model.insert_collection(dir_id, self.current_collection_type, file, '')
                doc = Document(os.path.join(folder, file))
                self.parser.run(doc, collection_id)
                self.model.db_connection.commit()

                continue
            if os.path.isdir(os.path.join(folder, file)):
                print("Folder {}".format(os.path.join(folder, file)))
                last_dir_id = self.model.insert_dir(dir_id, file)
                self.walk(os.path.join(folder, file), last_dir_id)
