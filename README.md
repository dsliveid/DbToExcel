# TableToStruct
表结构与Word文档之间互相转换工具

### 创建虚拟环境
##### venv目录通常是指一个虚拟环境（virtual environment）
python -m venv venv

#### 激活虚拟环境
##### 在Unix或MacOS上：
source venv/bin/activate
##### 在Windows上：
venv\Scripts\activate


### 安装依赖
##### 运行环境
链接数据库工具

pip install pyodbc

Excel导出依赖

pip install pandas openpyxl

打包为可执行文件需要

pip install pyinstaller

图片处理工具，图标需要

pip install Pillow

进度条工具，进度条需要

pip install tqdm

PgSql驱动，数据库连接需要

pip install psycopg2


### 打包命令
pyinstaller --onefile .\main.py

或者，带图标的打包

pyinstaller .\config.spec
