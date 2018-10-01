import os
from Model import FileModel, SqlModel
from Parsers.ParseCollection import Parse

from docx import Document
import json



class FolderWalker:
    def __init__(self, root_folder):
        self.root_folder = root_folder
        self.current_collection_type = None
        self.threads_data = list()

    def walk(self, folder, dir_id=None):
        for file in os.listdir(folder):
            if (file.startswith('~')): continue
            if file.endswith(".docx"):
                with open(os.path.join(folder, 'config.cfg')) as f:
                    config_file = json.load(f)
                    self.current_collection_type = config_file.get('type', 0)
                    self.current_prefix = config_file.get('id_prefix', '')
                    if self.current_collection_type == 2:
                        continue
                self.threads_data.append([os.path.join(folder, file), file, dir_id, self.current_collection_type, self.current_prefix])
                continue
            if os.path.isdir(os.path.join(folder, file)):
                last_dir_id = 0
                self.walk(os.path.join(folder, file), last_dir_id)
        return self.threads_data
