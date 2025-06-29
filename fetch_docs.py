#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
import ssl
import time
import json
import csv
import os

# 关闭 SSL 验证警告
requests.packages.urllib3.disable_warnings()
ssl._create_default_https_context = ssl._create_unverified_context

# 自动创建目录
output_dir = './gmbz_docs'
os.makedirs(output_dir, exist_ok=True)
csv_file_path = os.path.join(output_dir, 'gmbz_docs.csv')

def sanitize_filename(name: str):
    """去除非法文件名字符"""
    return ''.join(c for c in name if c not in r'\/:*?"<>|').strip()

def fetch_files():
    url = "http://www.gmbz.org.cn/main/normsearch.json"

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Language": "zh-CN,zh;q=0.8",
        "Accept-Encoding": "gzip, deflate",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "http://www.gmbz.org.cn",
        "Connection": "close"
    }

    data = {
        "draw": "1",
        "start": "0",
        "length": "500",  # 可调整，最大一次性获取数量
        "search[value]": "",
        "search[regex]": "false"
    }

    for i, name in enumerate([
        "NORM_ISO_ID", "NORM_NAME_C", "NORM_ZT_NAME", "NORM_FLAG_NAME",
        "NORM_NAME_E", "NORM_CO_NAME", "NORM_CA_NAME", "NORM_PUB_DATE",
        "NORM_IMP_DATE", "UP_GB_FLAG", "10"
    ]):
        data[f"columns[{i}][data]"] = name
        data[f"columns[{i}][name]"] = name
        data[f"columns[{i}][searchable]"] = "true"
        data[f"columns[{i}][orderable]"] = "true"
        data[f"columns[{i}][search][value]"] = ""
        data[f"columns[{i}][search][regex]"] = "false"

    data["order[0][column]"] = "0"
    data["order[0][dir]"] = "asc"

    print("[*] 正在请求标准列表...")

    try:
        response = requests.post(url, headers=headers, data=data, verify=False, timeout=10)
        response.encoding = 'utf-8'
    except requests.RequestException as e:
        print(f"[!] 请求失败：{e}")
        return

    if response.status_code != 200:
        print(f"[!] 请求失败，状态码：{response.status_code}")
        return

    try:
        content_json = response.json()
    except json.JSONDecodeError as e:
        print("[!] JSON解析失败，响应内容如下：")
        print(response.text[:500])
        return

    records = content_json.get('data', [])
    print(f"[*] 获取到 {len(records)} 条记录，正在写入 CSV 并下载文档...")

    with open(csv_file_path, "w", newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["index", "行标号", "标准中文名称", "类别", "状态", "发布日期", "实施日期", "文档下载链接"])

        for index, line in enumerate(records, start=1):
            norm_id = line.get('NORM_ID', '')
            name_c = line.get('NORM_NAME_C', '')
            zt = line.get('NORM_ZT_NAME', '--')
            flag = line.get('NORM_FLAG_NAME', '--')
            pub_date = line.get('NORM_PUB_DATE', '')
            imp_date = line.get('NORM_IMP_DATE', '')
            file_path = line.get('NORM_APP_ADDR', '')

            download_url = f"http://www.gmbz.org.cn/file/{file_path}" if file_path else ''
            writer.writerow([index, norm_id, name_c, zt, flag, pub_date, imp_date, download_url])

            # 下载 PDF 文件
            if file_path:
                safe_name = sanitize_filename(f"{norm_id} {name_c}.pdf")
                file_save_path = os.path.join(output_dir, safe_name)

                if not os.path.exists(file_save_path):
                    try:
                        r = requests.get(download_url, stream=True, verify=False, timeout=15)
                        with open(file_save_path, 'wb') as f:
                            for chunk in r.iter_content(chunk_size=1024):
                                if chunk:
                                    f.write(chunk)
                        print(f"[+] 下载成功：{safe_name}")
                    except Exception as e:
                        print(f"[!] 下载失败：{safe_name}，原因：{e}")
                else:
                    print(f"[=] 已存在跳过：{safe_name}")

    print(f"[*] 全部完成，CSV 和文档保存在：{output_dir}")

if __name__ == '__main__':
    fetch_files()
