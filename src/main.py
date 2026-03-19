import os
import sys
import pandas as pd
import logging
from tqdm import tqdm
import utils
import datetime

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def process_batches(column_data, query_condition, desc="整体进度"):
    """Process data in batches and return all results."""
    batch_size = 3000
    total_batches = (len(column_data) + batch_size - 1) // batch_size
    total_rows = len(column_data)
    logging.info(f"Processing {total_rows} rows in {total_batches} batches.")

    all_results = []
    keys = None
    pbar = tqdm(total=total_rows, desc=desc, unit="num", leave=True, position=0)
    for i in range(total_batches):
        start = i * batch_size
        end = min(start + batch_size, total_rows)
        batch_data = column_data[start:end]

        # Execute the query for the current batch
        results, batch_keys = utils.execute_query(batch_data, query_condition)
        if keys is None:
            keys = batch_keys
        all_results.extend(results)

        # Update progress bar
        mem = sys.getsizeof(batch_data) / (1024 ** 2)  # MB
        pbar.set_postfix_str(f"当前批次={i+1}/{total_batches}, 内存数据量={mem:.1f}MB, 总记录数={len(all_results)}", refresh=True)
        pbar.update(len(batch_data))

    pbar.close()
    return all_results, keys

def export_results(all_results, keys):
    """Export results to CSV and optionally to Excel if large."""
    if all_results:
        df = pd.DataFrame(all_results, columns=keys)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_filename = f"query_result_{timestamp}.csv"
        df.to_csv(csv_filename, index=False, encoding="utf-8-sig")
        logging.info(f"Results exported to {csv_filename}")
        
        if len(df) > 1048576:
            excel_filename = f"query_result_{timestamp}.xlsx"
            rows_per_sheet = 1000000
            num_sheets = (len(df) + rows_per_sheet - 1) // rows_per_sheet
            with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
                for i in range(num_sheets):
                    start = i * rows_per_sheet
                    end = min(start + rows_per_sheet, len(df))
                    sheet_df = df.iloc[start:end]
                    sheet_name = f"Sheet_{i+1}"
                    sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            logging.info(f"Large results also exported to {excel_filename} with {num_sheets} sheets.")
    else:
        logging.info("No results to export.")

# TODO: 增加批次写入和断点续传
def main():
    if len(sys.argv) < 4:
        logging.error("Usage: python main.py <file_path> <column_name_or_index> <query_sql_file> [has_header]")
        sys.exit(1)

    file_path = sys.argv[1]
    column_input = sys.argv[2]
    query_sql_file = sys.argv[3]
    has_header = True
    if len(sys.argv) >= 5:
        has_header = str(sys.argv[4]).strip().lower() in ("1", "true", "yes", "y")

    # Read query condition from the specified .sql file
    query_sql_path = os.path.join(os.path.dirname(__file__), '..', query_sql_file)
    if not os.path.isfile(query_sql_path):
        logging.error(f"Query SQL file not found: {query_sql_path}")
        sys.exit(1)
    with open(query_sql_path, 'r', encoding='utf-8') as f:
        query_condition = f.read().strip()

    if not os.path.isfile(file_path):
        logging.error(f"File not found: {file_path}")
        sys.exit(1)

    logging.info(f"Reading file: {file_path} (has_header={has_header})")
    row_threshold = 1000000
    if file_path.endswith('.xlsx'):
        sheets = pd.read_excel(file_path, sheet_name=None)
        total_rows = sum(len(df) for df in sheets.values())
        if total_rows > row_threshold:
            logging.warning(f"Excel file has {total_rows} rows, which exceeds {row_threshold}. 准备进行分批读取.")
        if len(sheets) > 1:
            logging.info(f"Found {len(sheets)} sheets, merging them.")
            if has_header:
                # Assume each sheet has header and headers are the same
                combined_data = pd.concat(sheets.values(), ignore_index=True)
            else:
                combined_data = pd.concat(sheets.values(), ignore_index=True)
            data = combined_data
        else:
            if has_header:
                data = pd.read_excel(file_path)
            else:
                data = pd.read_excel(file_path, header=None)
    elif file_path.endswith('.csv'):
        # Estimate total rows
        with open(file_path, 'r', encoding='utf-8') as f:
            total_rows = sum(1 for _ in f)
        if total_rows > row_threshold:
            logging.info(f"CSV file has {total_rows} rows, 准备进行分批读取.")
            chunk_size = 1000000
            chunks = pd.read_csv(file_path, chunksize=chunk_size, header=0 if has_header else None)
            all_results = []
            keys = None
            for chunk in chunks:
                # Process each chunk
                if not has_header:
                    column_data = chunk.iloc[:, 0]
                else:
                    # Determine column
                    if column_input.isdigit():
                        column_index = int(column_input) - 1
                        if column_index < 0 or column_index >= len(chunk.columns):
                            logging.error(f"Column index {column_input} is out of range in chunk.")
                            continue
                        column_data = chunk.iloc[:, column_index]
                    else:
                        if column_input not in chunk.columns:
                            logging.error(f"Column '{column_input}' not found in chunk.")
                            continue
                        column_data = chunk[column_input]
                
                # Batch process for this chunk
                results, batch_keys = process_batches(column_data, query_condition, desc="整体进度")
                if keys is None:
                    keys = batch_keys
                all_results.extend(results)
            
            # Export after all chunks
            export_results(all_results, keys)
            return  # Exit after chunk processing
        else:
            # Normal read for smaller CSV
            if has_header:
                data = pd.read_csv(file_path)
            else:
                data = pd.read_csv(file_path, header=None)
    else:
        logging.error("Unsupported file format. Please provide a .xlsx or .csv file.")
        sys.exit(1)

    # For Excel or small CSV, proceed as before
    if not has_header:
        column_data = data.iloc[:, 0]
    else:
        # Determine if the column input is a name or an index
        if column_input.isdigit():
            column_index = int(column_input) - 1  # Convert to zero-based index
            column_name = None
        else:
            column_name = column_input
            column_index = None

        # Determine the column to process
        if column_name:
            if column_name not in data.columns:
                logging.error(f"Column '{column_name}' not found in the data.")
                sys.exit(1)
            column_data = data[column_name]
        else:
            if column_index < 0 or column_index >= len(data.columns):
                logging.error(f"Column index {column_input} is out of range.")
                sys.exit(1)
            column_data = data.iloc[:, column_index]

    # Process data in batches
    all_results, keys = process_batches(column_data, query_condition)

    # Export results
    export_results(all_results, keys)

if __name__ == "__main__":
    main()