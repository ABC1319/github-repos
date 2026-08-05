#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub Repos Exporter v3.0
导出 GitHub 个人全部仓库（公开+私有）为 JSON / CSV / XLSX / HTML 单页导航
Fork & Modified from github-star: https://github.com/ABC1319/github-star

获取私有仓库说明：
- 必须提供 GitHub Token（classic），且勾选 repo 权限
- 使用 /user/repos API 获取当前认证用户的全部仓库（公开+私有）
- 不提供 Token 时，只能获取指定用户的公开仓库
"""

import requests
import json
import csv
import os
import sys
import time
import argparse
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

API_BASE = "https://api.github.com"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def get_headers(token):
    headers = {
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "github-repos-exporter/3.0"
    }
    if token:
        headers["Authorization"] = f"token {token}"
    return headers


def fetch_all_repos(username, token):
    """
    获取用户的全部仓库。
    如果提供了 Token，使用 /user/repos 获取当前认证用户的全部仓库（公开+私有）。
    如果没有 Token，使用 /users/{username}/repos 获取指定用户的公开仓库。
    """
    repos = []
    page = 1
    per_page = 100
    headers = get_headers(token)

    if token:
        # 使用 /user/repos 获取当前认证用户的全部仓库（包括私有）
        print(f"[1/4] 正在获取当前认证用户的全部仓库（公开+私有）...")
        url_template = f"{API_BASE}/user/repos?per_page={per_page}&page={{page}}&sort=updated&direction=desc&affiliation=owner,collaborator,organization_member"
    else:
        # 未提供 Token，只能获取指定用户的公开仓库
        print(f"[1/4] 正在获取用户 {username} 的公开仓库...")
        url_template = f"{API_BASE}/users/{username}/repos?per_page={per_page}&page={{page}}&sort=updated&direction=desc"

    while True:
        url = url_template.format(page=page)
        resp = requests.get(url, headers=headers, timeout=30)

        if resp.status_code == 401:
            print(f" 错误：Token 无效或已过期，请检查 GITHUB_TOKEN")
            return []
        if resp.status_code == 404:
            print(f" 错误：用户 '{username}' 不存在")
            return []
        if resp.status_code == 403:
            rate_limit_msg = ""
            try:
                reset_time = int(resp.headers.get("X-RateLimit-Reset", 0))
                if reset_time:
                    wait_min = int((reset_time - time.time()) / 60)
                    rate_limit_msg = f"，将在 {wait_min} 分钟后重置"
            except:
                pass
            print(f" 错误：API 速率限制{rate_limit_msg}。建议添加 GITHUB_TOKEN")
            return []

        resp.raise_for_status()
        data = resp.json()
        if not data:
            break

        repos.extend(data)
        print(f" 第 {page} 页: {len(data)} 个，累计 {len(repos)} 个")
        if len(data) < per_page:
            break
        page += 1
        time.sleep(0.3)

    # 去重（按 id）
    seen = set()
    unique_repos = []
    for r in repos:
        rid = r.get("id")
        if rid not in seen:
            seen.add(rid)
            unique_repos.append(r)

    print(f" 共获取 {len(unique_repos)} 个唯一仓库（去重前 {len(repos)} 个）\n")
    return unique_repos


def build_repo_data(repos):
    results = []
    total = len(repos)
    print(f"[2/4] 正在处理 {total} 个仓库...")
    for idx, repo in enumerate(repos, 1):
        item = {
            "序号": idx,
            "项目名称": repo.get("name", ""),
            "项目全名": repo.get("full_name", ""),
            "项目链接": repo.get("html_url", ""),
            "仓库描述": repo.get("description") or "",
            "主页网址": repo.get("homepage") or "",
            "项目语言": repo.get("language") or "",
            "星标数": repo.get("stargazers_count", 0),
            "Fork数": repo.get("forks_count", 0),
            "Watch数": repo.get("watchers_count", 0),
            "Open_Issues": repo.get("open_issues_count", 0),
            "默认分支": repo.get("default_branch", ""),
            "创建时间": repo.get("created_at", ""),
            "更新时间": repo.get("updated_at", ""),
            "推送时间": repo.get("pushed_at", ""),
            "仓库大小_KB": repo.get("size", 0),
            "是否为Fork": repo.get("fork", False),
            "是否私有": repo.get("private", False),
            "Topics": ", ".join(repo.get("topics", [])),
            "License": (repo.get("license") or {}).get("spdx_id", ""),
            "Owner": repo.get("owner", {}).get("login", ""),
            "Owner类型": repo.get("owner", {}).get("type", ""),
            "Owner头像": repo.get("owner", {}).get("avatar_url", ""),
        }
        results.append(item)
        if idx % 50 == 0 or idx == total:
            print(f" [{idx}/{total}] 已处理")
    print(f" 全部处理完成\n")
    return results


def export_json(data, filepath):
    os.makedirs(os.path.dirname(filepath) or '.', exist_ok=True)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f" JSON: {filepath}")


def export_csv(data, filepath):
    if not data:
        return
    os.makedirs(os.path.dirname(filepath) or '.', exist_ok=True)
    fieldnames = list(data[0].keys())
    csv_fieldnames = [f for f in fieldnames if f not in ["Owner头像"]]
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=csv_fieldnames)
        writer.writeheader()
        for row in data:
            writer.writerow({k: v for k, v in row.items() if k in csv_fieldnames})
    print(f" CSV: {filepath}")


def export_xlsx(data, filepath):
    if not HAS_OPENPYXL:
        print(" 未安装 openpyxl，跳过 XLSX")
        return
    if not data:
        return
    os.makedirs(os.path.dirname(filepath) or '.', exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "GitHub Repos"
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="24292F", end_color="24292F", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE")
    )
    fieldnames = list(data[0].keys())
    for col_idx, field in enumerate(fieldnames, 1):
        cell = ws.cell(row=1, column=col_idx, value=field)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    for row_idx, row in enumerate(data, 2):
        for col_idx, field in enumerate(fieldnames, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row.get(field, ""))
            cell.alignment = cell_align
            cell.border = thin_border
    col_widths = {
        "序号": 6, "项目名称": 25, "项目全名": 35, "项目链接": 45,
        "仓库描述": 50, "主页网址": 40, "项目语言": 12,
        "星标数": 10, "Fork数": 10, "Watch数": 10, "Open_Issues": 12,
        "默认分支": 12, "创建时间": 20, "更新时间": 20, "推送时间": 20,
        "仓库大小_KB": 14, "是否为Fork": 12, "是否私有": 12, "Topics": 40,
        "License": 20, "Owner": 18, "Owner类型": 12, "Owner头像": 45
    }
    for col_idx, field in enumerate(fieldnames, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(field, 20)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(filepath)
    print(f" XLSX: {filepath}")


def export_html(data, filepath, username):
    if not data:
        return
    os.makedirs(os.path.dirname(filepath) or '.', exist_ok=True)

    total_stars = sum(r.get("星标数", 0) for r in data)
    total_forks = sum(r.get("Fork数", 0) for r in data)
    private_count = sum(1 for r in data if r.get("是否私有"))
    lang_counts = {}
    for r in data:
        lang = r.get("项目语言") or "Unknown"
        lang_counts[lang] = lang_counts.get(lang, 0) + 1

    js_data = []
    for r in data:
        js_data.append({
            "name": r["项目名称"],
            "full_name": r["项目全名"],
            "html_url": r["项目链接"],
            "description": r["仓库描述"],
            "homepage": r["主页网址"],
            "language": r["项目语言"],
            "topics": r["Topics"],
            "stargazers_count": r["星标数"],
            "forks_count": r["Fork数"],
            "watchers_count": r["Watch数"],
            "open_issues_count": r["Open_Issues"],
            "license": r["License"],
            "created_at": r["创建时间"],
            "updated_at": r["更新时间"],
            "pushed_at": r["推送时间"],
            "size": r["仓库大小_KB"],
            "is_fork": r["是否为Fork"],
            "is_private": r["是否私有"],
            "owner_login": r["Owner"],
            "owner_avatar": r["Owner头像"],
            "default_branch": r["默认分支"],
        })
    js_json = json.dumps(js_data, ensure_ascii=False)
    gen_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    template_path = os.path.join(SCRIPT_DIR, "template.html")
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    html_content = (template
        .replace("{{USERNAME}}", username)
        .replace("{{TOTAL}}", str(len(data)))
        .replace("{{STARS}}", str(total_stars))
        .replace("{{FORKS}}", str(total_forks))
        .replace("{{PRIVATE}}", str(private_count))
        .replace("{{LANGS}}", str(len(lang_counts)))
        .replace("{{TIME}}", gen_time)
        .replace("{{JS_DATA}}", js_json))

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f" HTML: {filepath}")


def main():
    parser = argparse.ArgumentParser(description="导出 GitHub 个人全部仓库（公开+私有）")
    parser.add_argument("--output", "-o", default="./dist", help="输出目录 (默认: ./dist)")
    parser.add_argument("--username", "-u", default=os.environ.get("GITHUB_USERNAME"), help="GitHub 用户名（仅用于显示，获取私有仓库需 Token）")
    parser.add_argument("--token", "-t", default=os.environ.get("GITHUB_TOKEN"), help="GitHub Personal Access Token（classic，需勾选 repo 权限）")
    parser.add_argument("--formats", "-f", default="json,csv,xlsx,html", help="导出格式，逗号分隔")
    args = parser.parse_args()

    username = args.username or "unknown"
    token = args.token
    output_dir = args.output
    formats = [f.strip().lower() for f in args.formats.split(",")]

    if not token:
        print("⚠️  警告：未提供 GitHub Token，只能获取公开仓库。")
        print("    如需获取私有仓库，请提供带有 repo 权限的 Token。\n")

    print(f"输出目录: {output_dir}")
    print(f"导出格式: {formats}")
    print(f"GitHub Token: {'已提供 ✅' if token else '未提供 ❌'}\n")

    repos = fetch_all_repos(username, token)
    if not repos:
        print("未获取到任何仓库，退出。")
        sys.exit(1)

    data = build_repo_data(repos)

    print("[3/4] 正在导出文件...")
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"github_repos_{username}_{timestamp}"

    for fmt in formats:
        if fmt == "json":
            export_json(data, os.path.join(output_dir, f"{base_name}.json"))
            export_json(data, os.path.join(output_dir, "repos.json"))
        elif fmt == "csv":
            export_csv(data, os.path.join(output_dir, f"{base_name}.csv"))
            export_csv(data, os.path.join(output_dir, "repos.csv"))
        elif fmt == "xlsx":
            export_xlsx(data, os.path.join(output_dir, f"{base_name}.xlsx"))
            export_xlsx(data, os.path.join(output_dir, "repos.xlsx"))
        elif fmt == "html":
            export_html(data, os.path.join(output_dir, "index.html"), username)
        else:
            print(f" 未知格式: {fmt}")

    print("\n[4/4] 全部完成！")
    print(f" 输出目录: {output_dir}")
    for f in sorted(os.listdir(output_dir)):
        print(f" - {f}")


if __name__ == "__main__":
    main()
