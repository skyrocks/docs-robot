# docs-robot
这是一个Python写的自动生产数据库文档的脚本

### 环境
你需要在Python3环境下安装以下程序模块
```
pip install document
pip install pymysql
```

### 执行
运行app.py
```
python3 app.py
```

### 结果
生产数据库结构文档内容包括：<br>
>表名<br>
>序号，字段，名称，类型，默认值，主键，允许空值 

### 限制
只适用于mysql数据库