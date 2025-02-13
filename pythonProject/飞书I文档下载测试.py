import os
import re
import requests
import pandas as pd
from datetime import datetime

# 飞书 API 配置
APP_ID = "cli_a72e948d61d2900e"
APP_SECRET = "DiWuiC8pF7SjZOknpehdvgXWZBmUMann"
ACCESS_TOKEN_URL = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"


def get_desktop_path():
    """获取用户桌面路径"""
    return os.path.join(os.path.expanduser("~"), "Desktop")


def get_access_token():
    """获取访问令牌"""
    payload = {"app_id": APP_ID, "app_secret": APP_SECRET}
    try:
        response = requests.post(ACCESS_TOKEN_URL, json=payload, timeout=10)
        response.raise_for_status()
        return response.json().get("tenant_access_token")
    except Exception as e:
        print(f"❌ 获取Token失败: {str(e)}")
        return None


def parse_document_url(document_url):
    """解析多维表格URL"""
    pattern = r"/base/([a-zA-Z0-9_]+).*table=([a-zA-Z0-9_]+)"
    match = re.search(pattern, document_url)
    if not match:
        return None, None
    return match.group(1), match.group(2)


def get_app_name(access_token, base_token):
    """获取多维表格应用名称"""
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{base_token}"
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return response.json().get("data", {}).get("name", "未知文档")
    except Exception as e:
        print(f"⚠️ 获取文档名称失败: {str(e)}")
        return "飞书文档"


def get_table_data(access_token, base_token, table_id):
    """获取完整表格数据（自动分页）"""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{base_token}/tables/{table_id}/records"

    all_records = []
    page_token = ""
    while True:
        params = {"page_size": 100}
        if page_token:
            params["page_token"] = page_token

        try:
            response = requests.get(url, headers=headers, params=params, timeout=15)
            response.raise_for_status()
            data = response.json()

            if data.get("code") != 0:
                print(f"❌ API返回错误: {data.get('msg')}")
                return None

            all_records.extend(data.get("data", {}).get("items", []))

            if not data.get("data", {}).get("has_more"):
                break

            page_token = data.get("data", {}).get("page_token", "")
        except Exception as e:
            print(f"❌ 获取数据失败: {str(e)}")
            return None

    return {"data": {"items": all_records}}


def save_to_file(df, save_path, file_name, file_format):
    """保存数据到文件"""
    full_path = os.path.join(save_path, f"{file_name}.{file_format}")
    try:
        if file_format == "csv":
            df.to_csv(full_path, index=False, encoding="utf-8-sig")
        elif file_format == "xlsx":
            df.to_excel(full_path, index=False)
        print(f"✅ 文件已保存至：{full_path}")
        return True
    except Exception as e:
        print(f"❌ 保存文件失败: {str(e)}")
        return False


def main():
    try:
        # 获取访问令牌
        access_token = get_access_token()
        if not access_token:
            print("❌ 无法获取访问令牌，请检查APP_ID和APP_SECRET是否正确。")
            return

        # 获取用户输入
        document_url = input("请输入多维表格URL：").strip()
        if not document_url:
            print("❌ URL不能为空，请重新输入。")
            return

        base_token, table_id = parse_document_url(document_url)
        if not all([base_token, table_id]):
            print("❌ URL格式错误，请确认包含/base/和table参数。")
            return

        # 获取文档信息
        doc_name = get_app_name(access_token, base_token)
        file_name = f"{doc_name}_{datetime.now().strftime('%Y%m%d%H%M')}"  # 确保这里不包含文件扩展名

        # 获取表格数据
        print("⏳ 正在获取表格数据...")
        content = get_table_data(access_token, base_token, table_id)
        if not content:
            print("❌ 获取表格数据失败，请检查网络连接或API权限。")
            return

        # 处理数据
        df = pd.DataFrame([item["fields"] for item in content["data"]["items"]])
        if df.empty:
            print("❌ 未获取到有效数据。")
            return

        # 保存文件
        desktop_path = get_desktop_path()
        print("请选择导出格式：")
        print("1. Excel 文件 (.xlsx)")
        print("2. CSV 文件 (.csv)")
        choice = input("请输入数字选择（默认1）: ").strip() or "1"

        file_format = "xlsx" if choice == "1" else "csv"
        print(f"Saving file as {file_format} format...")  # 调试输出
        save_to_file(df, desktop_path, file_name, file_format)

    except Exception as e:
        print(f"❌ 程序执行过程中发生未知错误: {str(e)}")

if __name__ == "__main__":
    main()

