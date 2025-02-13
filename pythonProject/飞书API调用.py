import requests
import pandas as pd

# 飞书 API 配置
APP_ID = "cli_a72e948d61d2900e"
APP_SECRET = "DiWuiC8pF7SjZOknpehdvgXWZBmUMann"
ACCESS_TOKEN_URL = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"


def get_access_token():
    """获取访问令牌"""
    payload = {"app_id": APP_ID, "app_secret": APP_SECRET}
    response = requests.post(ACCESS_TOKEN_URL, json=payload)
    if response.status_code == 200:
        return response.json().get("tenant_access_token")
    else:
        print(f"获取Token失败: {response.text}")
        return None


def read_document_content(access_token, base_token, table_id):
    """读取多维表格数据"""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{base_token}/tables/{table_id}/records"
    response = requests.get(url, headers=headers)
    print("Request URL:", url)
    print("Response Status:", response.status_code)
    try:
        return response.json()
    except Exception as e:
        print(f"解析JSON失败: {str(e)}")
        return {}


def extract_table_data(content):
    """提取数据到DataFrame"""
    try:
        records = content["data"]["items"]
        data = []
        for record in records:
            fields = record["fields"]
            data.append(fields)
        return pd.DataFrame(data)
    except KeyError as e:
        print(f"KeyError: 缺失字段 {str(e)}")
        return None


def main():
    access_token = get_access_token()
    if not access_token:
        print("❌ 无法获取访问令牌")
        return

    document_url = input("请输入多维表格URL：").strip()
    try:
        parts = document_url.split('/base/')
        base_token = parts[1].split('?')[0]
        table_id = parts[1].split('table=')[1].split('&')[0]
    except IndexError:
        print("❌ URL格式错误")
        return

    content = read_document_content(access_token, base_token, table_id)
    print("API返回内容:", content)

    if content.get("code") == 0:
        df = extract_table_data(content)
        if df is not None:
            df.to_csv("output.csv", index=False)
            print("数据已保存至 output.csv")
    else:
        print(f"❌ 请求失败: {content.get('msg')}")


if __name__ == "__main__":
    main()
