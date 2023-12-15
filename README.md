# tony-sales-data

1. 安装 python 3.9
2. 安装依赖 `pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt`
3. 设置 processdata.py 里的以下参数:
   1. header_dict: 设置 平台名称对应的 缩写代码, 用于生成 json 文件的 header 数据
   2. orderline_setup: Sales Invoice 文件的 orderline 数据需要的 列名,以及对应的 excel 表头的名称
   3. scmline_setup: SCM 文件的 orderline 数据需要的 列名,以及对应的 excel 表头的名称
4. 运行 `python processdata.py`
