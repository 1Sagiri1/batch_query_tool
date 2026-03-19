import pandas as pd
import logging
from tqdm import tqdm
from urllib.parse import quote_plus
from sqlalchemy import create_engine, text
from config import DB_CONFIG

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def execute_query(batch_data, query_condition):
    """Execute a batch query using the provided batch data and query condition."""

    engine = create_engine(
        f"{DB_CONFIG['DIALECT']}+{DB_CONFIG['DRIVER']}://"
        f"{quote_plus(DB_CONFIG['USERNAME'])}:{quote_plus(DB_CONFIG['PASSWORD'])}@"
        f"{DB_CONFIG['HOST']}:{DB_CONFIG['PORT']}/"
        f"{DB_CONFIG['DATABASE']}"
    )
    # Prepare the values from batch_data (assuming it's a Series)
    clean_data = batch_data.dropna().unique()
    
    # 关键：过滤后为空则直接返回，避免 IN () 语法错误
    if len(clean_data) == 0:
        return [], []
    
    values = ', '.join(f"'{value}'" for value in clean_data)
    
    # Format the query condition with the values
    query = query_condition.format(values)
    
    # logging.info(f"Executing query: {query}")
    
    with engine.connect() as connection:
        result = connection.execute(text(query))
        results = result.fetchall()
        keys = result.keys()
        # tqdm.write(f"Query executed successfully, returned {len(results)} rows.")
    
    return results, keys

def log_query_results(results, limit=5):
    for i, result in enumerate(results):
        if i >= limit:
            break
        logging.info(result)