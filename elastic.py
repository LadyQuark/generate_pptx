import os
import pprint
from dotenv import load_dotenv, find_dotenv
from elasticsearch import Elasticsearch
from functools import wraps

pp = pprint.PrettyPrinter(depth=6)  
load_dotenv(find_dotenv())
# Connect to Elasic Cloud
ELASTIC_CLOUD_ID = os.getenv('ELASTIC_CLOUD_ID')
ELASTIC_USER = os.getenv('ELASTIC_USER')
ELASTIC_PASSWORD = os.getenv('ELASTIC_PASSWORD')
REQUEST_TIMEOUT = 900
MAX_RETRIES = 10
BATCH = 1000

if ELASTIC_CLOUD_ID:
    es = Elasticsearch(
        cloud_id=ELASTIC_CLOUD_ID,
        basic_auth=(ELASTIC_USER, ELASTIC_PASSWORD),
        max_retries=MAX_RETRIES, retry_on_timeout=True,
        request_timeout=REQUEST_TIMEOUT
    )
else:
    es = Elasticsearch(
        "http://localhost:9200",
        max_retries=MAX_RETRIES, retry_on_timeout=True,
        request_timeout=REQUEST_TIMEOUT
    )
INDEX = "docs"

def check_connection(f):
    @wraps(f)
    def decorated_func(*args, **kwargs):
        if es.ping():
            return f(*args, **kwargs)
        else:
            raise Exception("Could not connect to ElasticSearch")
    return decorated_func

MAPPINGS = {
    "properties" : {
        "user_id" : {
            "type" : "keyword",
            "index" : "true" 
        },
        "title": {
            "type": "text"
        },        
        "content": {
            "type": "text"
        },        
        "slide_id" : {
            "type" : "keyword",
            "index" : "false" 
        },
        "slide_index" : {
            "type" : "integer",
        },
        "virtualFileName" : {
            "type" : "keyword",
            "index" : "false" 
        },
        "originalFileName" : {
            "type" : "keyword",
            "index" : "false" 
        },
        "root" : {
            "type" : "keyword",
            "index" : "true" 
        },
    }
}