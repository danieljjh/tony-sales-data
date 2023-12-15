import os
import pandas as pd
import json
import datetime as dt

## 这里设置平台对应的编码
header_dict = [
    {
        "Description": "抖音",
        "Short Form": "DY",
        "B2C": "B2C000002"
    },
    {
        "Description": "天猫",
        "Short Form": "TM",
        "B2C": "B2C000003"
    },
    {
        "Description": "京东",
        "Short Form": "JD",
        "B2C": "B2C000004"
    },
    {
        "Description": "VIP",
        "Short Form": "VIP",
        "B2C": "B2C000005"
    },
    {
        "Description": "小红书",
        "Short Form": "RED",
        "B2C": "B2C000006"
    },
    {
        "Description": "爱库存",
        "Short Form": "Aikucun",
        "B2C": "B2C000006"
    },
    {
        "Description": "天猫Ipanema",
        "Short Form": "TM Ipanema",
        "B2C": "B2C-000006"
    }
]

orderline_setup =  {
            "Line Discount Amount": "*:平摊优惠",
            "Quantity": "*:数量",
            "Online Shop Order No.": "*:订单号",
            "No.": "*:关联销售物料编码",
            "Unit Price": "*:销售总价"
        }
scmline_setup = {
            "Line Discount Amount": "0.00",
            "Quantity": "1",
            "Online Shop Order No.": "1980139836701806876",
            "No.": "SERVICE ITEM-S/P",
            "Unit Price": "420.20"
        },

def gen_sales_data(df):
    """处理销售数据"""
    # 整理需要从 excel 里取的列
    line_cols = []
    line_rename = {}
    for k, v in orderline_setup.items():
        line_rename[v]=k
        line_cols.append(v)
    # 下面是生成数据
    df.loc[:, 'PostDay'] = df["发货日期"].apply(lambda x:x[:10].replace("-",""))
    for grp, rows in df[:100].sort_values(["PostDay","平台"], ascending=[1,1]).groupby("PostDay"):
        # print(grp)
        for pt, pt_rows in rows.groupby("平台"):
            shop_info = next((x for x in header_dict if x["Description"] == pt),  {
                "Description": "NA",
                "Short Form": "NA",
                "B2C": "B2C-00000A"
            })
            lines = pt_rows[line_cols]
            new_lines = lines.rename(columns=line_rename )
            data = {
                "Posting Date": grp,
                "Document Type": "Sales",
                "Document No.": f"{shop_info['Short Form']}_SI_{grp}",
                "Sell-to Customer No.": shop_info["B2C"],
                "orderLine": new_lines.to_dict("records")
            }
            file_path = f"./out/{shop_info['Short Form']}_SI_{grp}.B2C"
            with open(file_path, 'w') as json_file: 
                json.dump(data, json_file, indent=4)
            print(f"successful gen file {file_path}")

def gen_scm_data(df):
    """处理销售数据"""
    # 整理需要从 excel 里取的列
    line_cols = []
    line_rename = {}
    for k, v in scmline_setup.items():
        line_rename[v]=k
        line_cols.append(v)
    # 下面是生成数据
    df.loc[:, 'PostDay'] = df["发货日期"].apply(lambda x:x[:10].replace("-",""))
    for grp, rows in df[:100].sort_values(["PostDay","平台"], ascending=[1,1]).groupby("PostDay"):
        # print(grp)
        for pt, pt_rows in rows.groupby("平台"):
            shop_info = next((x for x in header_dict if x["Description"] == pt),  {
                "Description": "NA",
                "Short Form": "NA",
                "B2C": "B2C-00000A"
            })
            lines = pt_rows[line_cols]
            new_lines = lines.rename(columns=line_rename )
            data = {
                "Posting Date": grp,
                "Document Type": "Sales",
                "Document No.": f"{shop_info['Short Form']}_SCM_{grp}",
                "Sell-to Customer No.": shop_info["B2C"],
                "orderLine": new_lines.to_dict("records")
            }
            file_path = f"./out/{shop_info['Short Form']}_SI_{grp}.B2C"
            with open(file_path, 'w') as json_file: 
                json.dump(data, json_file, indent=4)
            print(f"successful gen file {file_path}")

def main():
    # Get the current directory
    current_dir = os.getcwd()



    # Get a list of all Excel files in the current directory
    excel_files = [file for file in os.listdir(current_dir) if file.endswith('.xlsx')]

    if len(excel_files) > 0:

        # Iterate over each Excel file and read it into a pandas DataFrame
        dataframes = []
        dataframes_scm = []
        for file in excel_files:
            excel_path = os.path.join(current_dir, file)
            df = pd.read_excel(excel_path, sheet_name=1)
            df_scm = pd.read_excel(excel_path, sheet_name=2)

            dataframes.append(df)
            dataframes_scm.append(df_scm)

        # Concatenate all dataframes into a single dataframe
        combined_df = pd.concat(dataframes)
        combined_df_scm = pd.concat(dataframes_scm)

        # Export the combined dataframe as a JSON file

        gen_sales_data(combined_df)
        try:
            gen_scm_data(combined_df_scm)
        except:
            print("SCM 文件错误")
    else:
        print("当前文件夹下没有找到 excel 文件")


if __name__ == '__main__':
    main()
