from multiprocessing.pool import ThreadPool

from docx import Document

from FolderWalker import FolderWalker
from Model import FileModel, SqlModel
from Parsers.ParseCollection import Parse

folder_walker = FolderWalker(r'крым')
thread_data = (folder_walker.walk(r'крым'))
model = SqlModel('result.sqlite')
def process(args):
    print("Start process with filename: {}\n".format(args[0]))
    collection_id = model.insert_collection(args[2], args[3], args[1], '')
    parser = Parse(model)
    doc = Document(args[0])
    parser.run(doc, collection_id, id_prefix=args[4], collection_type=args[3])
    model.commit()

for data in thread_data:
    process(data)

def calculate_parallel(args, threads = 1):
    pool = ThreadPool(threads)
    pool.map(process, args)
    pool.close()
    pool.join()

# calculate_parallel(thread_data)
