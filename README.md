# Excel Batch Query Tool

## 项目简介
该项目是一个用于批处理查询的工具，支持通过拖拽 `.xlsx` 或 `.csv` 文件到脚本上，并根据用户输入的列名或列序列进行数据查询。查询结果将通过命令行输出，并提供详细的日志记录和进度条显示。

## 文件结构
```
excel_batch_query
├── .venv                 # uv虚拟环境
├── src
│   ├── main.py           # 应用程序的入口点
│   ├── config.py         # 数据库连接配置
│   └── utils.py          # 实用函数
├── requirements.txt      # 项目依赖
└── README.md             # 项目文档
```

## 功能特性
- 支持.xlsx和.csv文件格式。
- 通过命令行参数指定列名或列索引（从1开始）。
- 支持批处理查询，默认每批3000行。
- 数据库连接参数在config.py中预留。
- 查询条件通过命令行输入，支持格式化字符串（如WHERE field IN ({})）。
- 详细日志输出，包括进度条。
- 查询结果导出为带时间戳的CSV文件。

## 使用前
1. 确保已安装uv（如果未安装，运行pip install uv）。

2. 导航到项目文件夹

3. 创建虚拟环境并激活：
   ```
   uv venv
   // win
   .\.venv\Scripts\activate
   // macos
   source ./.venv/bin/activate
   ```

4. 安装所需的依赖：
   ```
   uv pip install -r requirements.txt
   ```

## 数据库配置
运行前请先复制 src/config.example.py 并重命名为 config.py，
然后编辑src/config.py文件，设置数据库连接参数：

```
DB_CONFIG = {
    'DIALECT': 'mysql',
    'DRIVER': 'pymysql',
    'USERNAME': 'your_username',
    'PASSWORD': 'your_password',
    'HOST': 'your_host',
    'PORT': '3306',
    'DATABASE': 'your_database'
}
```

## 使用示例
1. 将 `.xlsx` 或 `.csv` 文件拖拽到 `src/main.py` 上。
2. 在命令行中输入要查询的列名或列序列（从1开始），并提供查询语句文件。例如：
   ```
   python src\main.py <file_path> <column_name_or_index> <query_sql_file> [has_header]
   python src\main.py "202602清关数据拉取.xlsx" 1 "query1.sql"
   ```

- `<file_path>`: .xlsx或.csv文件的路径（支持拖拽）。
- `<column_name_or_index>`: 列名或列索引（从1开始）。
- `<query_sql_file>`: 查询语句文件，如`"SELECT * FROM table WHERE field IN ({})"`。
- `[has_header]`: 可选，文件是否包含列名（1/true/yes/y表示是，默认是）。

结果将导出为query_result_YYYYMMDD_HHMMSS.csv。

## 注意事项
- 确保数据库连接参数正确。
- 查询条件中的{}将被替换为批处理的值列表。
- 日志输出到控制台，包括查询进度。