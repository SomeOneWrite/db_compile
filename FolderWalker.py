import os
from Model import FileModel, SqlModel
from Parsers.ParseCollection import Parse

from docx import Document
import json



class FolderWalker:
    def __init__(self, root_folder, model):
        self.root_folder = root_folder
        self.current_collection_type = None
        self.threads_data = list()
        self.last_dir_id = None
        self.model = model
        self.filenames = None

    def walk(self, folder, dir_id=None):
        for file in os.listdir(folder):
            if (file.startswith('~')): continue
            if os.path.exists(os.path.join(folder, 'config.cfg')):
                with open(os.path.join(folder, 'config.cfg'), encoding='UTF-8') as f:
                    config_file = json.load(f)
                    self.current_collection_type = config_file.get('type', 0)
                    self.current_prefix = config_file.get('id_prefix', '')
                    self.filenames = config_file.get('files')
            if self.filenames:
                for file_2 in self.filenames:
                    self.threads_data.append(
                        [os.path.join(folder, file_2), file_2, dir_id, self.current_collection_type,
                         self.current_prefix])
                self.filenames = None
                return
            if file.endswith(".docx"):
                if not self.filenames:
                    self.threads_data.append([os.path.join(folder, file), file, dir_id, self.current_collection_type, self.current_prefix])
                continue
            if os.path.isdir(os.path.join(folder, file)):
                self.last_dir_id = self.model.insert_dir(dir_id, file)
                self.walk(os.path.join(folder, file), self.last_dir_id)
        return self.threads_data
