import os
import Model

class FolderWalker:
    def __init__(self, root_folder):
        self.model = Model.Model("result.sqlite")
        self.root_folder = root_folder
        self.walk(root_folder)
        pass

    def walk(self, folder, dir_id = None):
        for file in os.listdir(folder):
            if os.path.isdir(os.path.join(folder, file)):
                print("Folder {}".format(os.path.join(folder, file)))
                last_dir_id = 0#self.model.insert_dir(file, dir_id)
                # last_parent_id = parent_id
                # parent_id = get_last_insert_id('dirs')
                self.walk(os.path.join(folder, file), last_dir_id)
                continue
            if file.endswith(".docx"):
                print("File {}".format(file))
                self.model.parse_file(os.path.join(folder, file))

