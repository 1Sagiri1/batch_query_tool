# 本地浏览器版批量查询工具

## 说明
项目已升级为“前端页面 + 本地数据库桥接 API”模式:
- 页面入口: `index.html`
- 本地桥接服务: `python src/bridge_server.py`

主要能力:

1. 多数据库/账号配置管理:
- 在页面填写配置并保存到 `localStorage`。
- 支持导入多个 JSON 配置文件。
- 支持标签页切换配置。

2. 文件导入与预览:
- 支持拖拽 `.csv/.xlsx/.xls`。
- 多 Sheet 自动转换为 CSV 处理。
- 大文件自动走 DuckDB（浏览器内）。
- 导入后可预览前 10 行，鼠标点击表头选批处理列。
- 支持“是否包含标题行”单选，避免漏查第一行。

3. 批处理查询:
- 支持多个任务并行执行。
- 每个任务可独立选择配置、SQL、批大小、导出格式。
- SQL 支持 `{}` 或 `{{values}}` 占位符，用于批量 `IN (...)` 替换。

4. 结果导出:
- 单任务导出 CSV 或 XLSX（超大结果自动多 Sheet）。
- 所有任务结果一键导出为 XLSX 多 Sheet。

## 使用方式
1. 安装依赖:
```bash
pip install -r requirements.txt
```
2. 双击 `start_app.bat`。
3. 会自动启动桥接服务并在“默认浏览器”打开网页。
4. 查询完成后回到启动窗口按 Enter，即可自动关闭桥接服务。
5. 其中 `start_app.bat` 是双击入口，`start_app.ps1` 负责进程生命周期控制。
5. 在“数据库/账号配置”里填写并测试连接。
6. 拖拽导入数据文件，选择批处理列，创建任务并运行。
7. 任务完成后导出 CSV 或 XLSX。

## 配置文件格式示例
```json
[
  {
    "name": "生产库A",
    "host": "127.0.0.1",
    "port": "3306",
    "username": "root",
    "password": "******",
    "database": "demo"
  }
]
```

## 兼容性提示
- 文件列值会分批替换到 SQL 的 `{}` / `{{values}}`，SQL 在真实数据库执行。
- `dialect/driver` 固定默认值为 `mysql+pymysql`，无需页面配置。
- DuckDB 仅用于浏览器内文件处理（提取去重值），不参与业务 SQL 查询。
- 若直接双击 `index.html`，无法自动管理本地桥接进程；请使用 `start_app.bat`。
