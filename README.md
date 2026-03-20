# 批量查询工作台（浏览器 + 本地桥接服务）

## 项目说明
本项目采用“浏览器页面 + 本地 Python 桥接服务”的方式运行：

- 前端页面：`index.html`
- 后端桥接服务：`src/bridge_server.py`
- 推荐启动入口：`start_app.bat`（自动启动服务并打开网页）

## 当前功能
1. 配置管理
- 支持多个数据库配置（名称、host、port、username、password、database）
- 支持配置导入/导出（JSON）
- 配置保存在浏览器 `localStorage`
- 可在页面内执行“连接测试（SELECT 1）”

2. 多文件导入与管理
- 支持连续导入多个 `.csv/.xlsx/.xls`
- 左侧“备选文件列表”显示已导入文件
- 每个文件独立保存：
  - 是否包含标题行
  - 选中的批处理列索引
  - 预览 sheet
- 支持删除误导入文件（释放占用，并清空相关任务绑定）

3. 任务并行执行
- 支持多个任务并行
- 每个任务可独立配置：
  - 绑定的数据库配置
  - 绑定的导入文件
  - SQL
  - 批大小
  - 导出格式
- SQL 支持 `{}` 或 `{{values}}` 占位符（用于批量 `IN (...)`）

4. 进度与中止
- 任务级进度 + 全局进度
- 显示耗时、ETA、批次、行数等指标
- 支持中止任务，中止后耗时和 ETA 冻结

5. 导出
- 单任务导出：CSV / XLSX
- 全部任务导出：多 sheet XLSX
- 文件命名：`任务名_时间戳`
- CSV 编码支持：
  - 自动（推荐，按系统适配）
  - UTF-8
  - UTF-8 BOM（Windows Excel 兼容）
- XLSX 导出默认走浏览器下载目录（不弹手动选路径）

## 快速开始
1. 安装依赖
```bash
pip install -r requirements.txt
```

2. 启动应用（推荐）
- 双击 `start_app.bat`
- 脚本会：
  - 启动桥接服务
  - 用浏览器打开页面
  - 等你回到启动窗口按 `Enter` 后关闭服务

3. 页面操作
- 在“数据库/账号配置”中填写并测试连接
- 导入一个或多个文件，逐个选择批处理列
- 新建任务并选择“绑定文件”
- 运行并导出结果

## 可选手动启动
如果不使用 `start_app.bat`，也可以手动：

1. 启动服务
```bash
python src/bridge_server.py
```
2. 打开 `index.html`

## 配置 JSON 示例
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

## 技术说明
- 业务 SQL 在真实数据库执行（通过本地桥接服务）
- DuckDB 仅用于浏览器端文件处理（提取去重值/预览辅助）
- `dialect/driver` 已固定默认 `mysql+pymysql`，页面无需配置
