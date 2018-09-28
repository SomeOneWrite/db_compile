import os
from Model import FileModel, SqlModel
from Parsers.ParseCollection import Parse


from threading import Thread
from docx import Document
import json

class CollectionThread(Thread):

    def __init__(self, filename, doc_filename, dir_id, collection_type):
        self.filename = filename
        self.doc_filename = doc_filename
        self.dir_id = dir_id
        self.collection_type = collection_type
        Thread.__init__(self)

    def run(self):
        model = FileModel(self.filename + '.json')
        collection_id = model.insert_collection(self.dir_id, self.collection_type, self.doc_filename, '')
        model.commit()
        parser = Parse(model)
        doc = Document(self.filename)
        parser.run(doc, collection_id)
        print('Start process')
        model.commit()

class FolderWalker:
    def __init__(self, root_folder):
        self.root_folder = root_folder
        self.walk(root_folder)
        self.current_collection_type = None
        self.threads = {}

    def walk(self, folder, dir_id=None):
        for file in os.listdir(folder):
            if (file.startswith('~')): continue
            if file.endswith(".docx"):
                with open(os.path.join(folder, 'config.cfg')) as f:
                    config_file = json.load(f)
                    self.current_collection_type = config_file['type']
                    print('collection type: {}'.format(self.current_collection_type))
                if self.current_collection_type != 2:
                    continue
                print("File {}".format(file))
                if len(self.threads) < 5:
                    my_thread = CollectionThread(os.path.join(folder, file), file, dir_id, self.current_collection_type)
                    my_thread.start()
                    self.threads[file] = my_thread
                continue
            if os.path.isdir(os.path.join(folder, file)):
                print("Folder {}".format(os.path.join(folder, file)))
                last_dir_id = 0
                self.walk(os.path.join(folder, file), last_dir_id)

