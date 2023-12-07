import pprint
from elasticsearch.helpers import bulk, BulkIndexError
from elasticsearch.exceptions import BadRequestError
from pathlib import Path
from tqdm import tqdm

from elastic import es, check_connection, INDEX, MAPPINGS, BATCH, REQUEST_TIMEOUT, MAX_RETRIES
from presentationmanager import PresentationManager

pp = pprint.PrettyPrinter(depth=6)  
TEST_USER_ID = "65700cee327beccab31fc13b"

@check_connection
def delete_create_index(index=INDEX):
    
    print("Deleting index:", index)
    es.options(ignore_status=404).indices.delete(
        index=index)
    print("Creating index:", index)
    es.indices.create(
        index=index,
        mappings=MAPPINGS
        )

@check_connection
def index_batch(docs, index=INDEX):
    """ Index documents in `index` in bulk in batches of size `BATCH`"""
    requests = []
    
    # Make list of requests
    total = len(docs)
    pbar = tqdm(docs)
    for doc in pbar:
        pbar.set_description("index")     
        # Prepare requests
        request = {}
        request = doc
        request["_op_type"] = "index"
        request["_index"] = index
        requests.append(request)
    
    success = 0
    errors = []
    # Index docs in batches of size BATCH
    for batch_request in chunks(requests, n=BATCH):
        try:
            count, e = bulk(client=es.options(
                    request_timeout=REQUEST_TIMEOUT,
                    max_retries=MAX_RETRIES, retry_on_timeout=True,
                    ), actions=batch_request, request_timeout=REQUEST_TIMEOUT)
        
        except BulkIndexError as e:
            # Print errors in detail
            print(e,  e.with_traceback)
            for item in e.errors:
                for key in item['index']:
                    if key != "data":
                        print(item['index'][key])
                print("\n")
            # Set number of successfully indexed documents        
            count = total - len(e.errors)
            errors.extend(e.errors)
        
        # Update number of indexed docs
        success += count

    return success, errors

def chunks(data, n):
    """ Generates chunks of given list """
    for i in range(0, len(data), n):
        yield data[i:i + n]

def get_all_docs(folder_path):
    folder = Path(folder_path)
    docs = []
    for filepath in folder.glob('*.pptx'):
        ppt = PresentationManager(filepath)
        filepath_str = str(filepath)
        results = ppt.extract_all_text()
        for result in results:
            result.update({
                "user_id": TEST_USER_ID,
                "root": folder_path,
                "virtualFileName": filepath.name,
                "originalFileName": filepath.name,                
            })
        docs.extend(results)
    return docs


@check_connection
def search_in_index(query,index=INDEX, size=10):
    from_index = 0

    try:
        resp = es.search(
                    index=index,
                    size=size,
                    from_=from_index,
                    query={"bool": {
                        "must": [
                            {"match" : {
                                "content": {
                                    "query": query,
                                    "fuzziness": "AUTO"
                                }                        
                            }},
                            {"term": {
                                "user_id": TEST_USER_ID
                            }}                                    
                        ]
                    }},
                    highlight={"fields": {
                        "content": {}
                        }},
                )
    except BadRequestError as e:
        print(f"{e} at {index}")
        return None
    
    return resp['hits']


def _strip_document(data):
    FIELDS = [
        "user_id",
        "title",
        "content",    
        "slide_id",
        "slide_index",
        "virtualFileName",
        "originalFileName",
        "root"
    ]
    OPTIONAL_FIELDS = [

    ]  
    try:
        doc = {
            key: data[key] for key in FIELDS
        }
    except KeyError as e:
        raise Exception("Document missing field:", e)
    
    for key in OPTIONAL_FIELDS:
        doc.update({
            key: data[key]
        })

    return doc 



@check_connection
def print_all_docs(index=INDEX, start_index=0, size=10):
    while True:
        resp = es.search(
            index=index, 
            from_=start_index,
            query={"match_all" : {}}, 
            size=10)
        total = resp['hits']['total']['value']
        print_results(resp['hits'], show_highlights=False) 
        start_index += size
        if start_index >= total:
            break


def print_results(hits, show_highlights=True):
    print("TOTAL:", hits['total']['value'])

    results = [item for item in hits['hits']]    
    
    column_name = "SEARCH RESULTS" if show_highlights else "TITLE/FIRST LINE"
    print("{:<35} {:<10s} {:s}".format("PRESENTATION", "SLIDE NO", column_name))  
    print("-" * 80)
    
    for r in results:
        column = ""
        if show_highlights:
            for highlights in r['highlight']['content']:
                lines = highlights.split("\n")
                for line in lines:
                    if "<em>" in line:
                        column = line
                        break
                if column:
                    break
        source = r['_source']
        file = source["originalFileName"]
        i = str(source["slide_index"] + 1)
        if not column:
            first_line = source["content"].split("\n", maxsplit=1)[0]
            column = source["title"] if source["title"] != "" else first_line[:50]
        print("{:<35} {:<10s} {:s}".format(file, i, column))    


