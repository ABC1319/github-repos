# GitHub Repos Exporter 🚀
![Export Status](https://github.com/ABC1319/github-repos/actions/workflows/export.yml/badge.svg)

> Fork & Modified from [github-star](https://github.com/ABC1319/github-star)  
> 将 GitHub 个人仓库导出为 JSON / CSV / XLSX / HTML 单页导航，支持自动部署到 GitHub Pages。



## ✨ 功能特性

- **🔍 实时搜索** — 支持搜索项目名称、描述、Topics、Owner、语言
- **✨ 关键词高亮** — 匹配的关键词自动高亮显示
- **🎛️ 多维度排序** — 默认 / 星标数 / 更新时间 / 名称 / 创建时间
- **🏷️ 语言筛选** — 按编程语言一键筛选
- **🌗 主题切换** — 白天/黑夜模式，自动保存偏好
- **🔝 一键置顶** — 滚动后显示回到顶部按钮
- **⌨️ 快捷键** — `Ctrl+K` 搜索，`Esc` 清空
- **📱 响应式** — 完美适配手机、平板、桌面
- **🤖 自动更新** — GitHub Actions 定时自动导出并部署

## 📦 安装依赖

```bash
pip install requests openpyxl
```

## 🚀 本地使用

### 方式 1：环境变量

```bash
export GITHUB_USERNAME="yourname"
export GITHUB_TOKEN="ghp_xxx"  # 可选，建议填写
python export.py
```

### 方式 2：命令行参数

```bash
python export.py --user yourname --token ghp_xxx --output ./dist
```

### 方式 3：指定格式

```bash
# 只导出 HTML
python export.py --user yourname --formats html

# 导出全部格式（默认）
python export.py --user yourname --formats json,csv,xlsx,html
```

### 命令行参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--user` | GitHub 用户名 | `GITHUB_USERNAME` 环境变量 |
| `--token` | GitHub Personal Access Token | `GITHUB_TOKEN` 环境变量 |
| `--output` | 输出目录 | `./dist` |
| `--formats` | 导出格式，逗号分隔 | `json,csv,xlsx,html` |

## 🔑 获取 GitHub Token

1. 打开 GitHub Settings → Developer settings → Personal access tokens → Tokens (classic)
2. 点击 **Generate new token (classic)**
3. 勾选 `repo` 和 `read:user` 权限
4. 点击 **Generate token**
5. 复制 Token（格式如 `ghp_xxxxxxxxxxxx`）

> **注意**：Token 仅用于提高 API 速率限制（5000次/小时 vs 60次/小时），建议填写。

## 🌐 自动部署到 GitHub Pages

本项目支持 GitHub Actions 自动定时导出并部署到 GitHub Pages。

### 配置步骤

1. **Fork 本仓库** 或创建新仓库，将 `export.py` 放入仓库根目录

2. **设置 Secrets**
   - 进入仓库 **Settings → Secrets and variables → Actions**
   - 点击 **New repository secret**
   - 添加以下 Secrets：
     - `GH_USERNAME`：你的 GitHub 用户名
     - `GH_TOKEN`：你的 GitHub Personal Access Token

3. **启用 GitHub Pages**
   - 进入仓库 **Settings → Pages**
   - **Source** 选择 **GitHub Actions**

4. **完成！**
   - 工作流会自动运行（每周一 00:00 UTC）
   - 也可手动触发：**Actions → Export GitHub Repos → Run workflow**
   - 部署完成后访问 `https://yourname.github.io/repo-name/`

### 工作流说明

`.github/workflows/export.yml` 包含以下配置：

- **定时触发**：每周一 00:00 UTC 自动运行
- **手动触发**：支持在 Actions 页面手动点击运行
- **自动部署**：生成 HTML 后自动推送到 GitHub Pages

### 修改定时规则

编辑 `.github/workflows/export.yml`：

```yaml
on:
  schedule:
    - cron: '0 0 * * 1'  # 每周一 00:00 UTC
```

Cron 表达式参考：

- `0 0 * * *` — 每天
- `0 0 * * 1` — 每周一
- `0 0 1 * *` — 每月1号

## 📁 输出文件

| 文件 | 说明 |
|------|------|
| `index.html` | 单页导航（搜索/排序/筛选/主题切换） |
| `github_repos_*.json` | 完整数据 |
| `github_repos_*.csv` | 表格数据 |
| `github_repos_*.xlsx` | Excel 表格（带样式） |

## 📄 数据字段

- 基本信息：项目名称、全名、链接
- 描述信息：仓库描述、主页网址
- 统计信息：星标数、Fork 数、Watch 数、Issues
- 分类信息：项目语言、Topics、License
- 时间信息：创建时间、更新时间、推送时间
- 其他：仓库大小、是否为 Fork、Owner 信息

## 📝 License

MIT
