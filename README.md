# 快乐8 开奖数据 · 自动更新 + Web 展示

从**中国福利彩票官网**自动抓取快乐8开奖数据，每天定时更新 JSON / Excel，并通过 GitHub Pages 提供一个静态 Web 页面查看开奖记录与号码统计。

## 目录结构

```
kl8-lottery/
├── fetch_kl8.py                  # 抓取 + 增量合并脚本
├── 快乐8开奖数据.json             # 结构化数据（web 页面数据源）
├── 快乐8开奖数据.xlsx             # Excel 表格（明细 + 频次）
├── index.html                    # Web 端页面
└── .github/workflows/update-kl8.yml   # 每日定时更新 + 部署 Pages
```

## 快速开始

### 1. 创建仓库并推送

```bash
git init
git add .
git commit -m "init: kl8 lottery data tracker"
git branch -M main
git remote add origin https://github.com/<你的用户名>/kl8-lottery.git
git push -u origin main
```

### 2. 开启 GitHub Pages

仓库 **Settings → Pages → Build and deployment**：

- Source 选择 **GitHub Actions**

之后每次 workflow 运行完都会自动部署，站点地址形如
`https://<你的用户名>.github.io/kl8-lottery/`。

> 也可以选择 Deploy from a branch（main / root），但那样页面不会随数据更新自动重新部署。

### 3. 允许 Actions 写入仓库

**Settings → Actions → General → Workflow permissions**，勾选
**Read and write permissions**（脚本需要把新数据 commit 回仓库）。

## 定时更新

`.github/workflows/update-kl8.yml`：

- 每天 **北京时间 22:10**（UTC 14:10，快乐8 每日 21:30 开奖后）自动运行
- 抓取官网最新数据，只合并缺失期号，重建 JSON 与 Excel
- 有新数据 → 自动 commit + push，并重新部署 Web 页面
- 没有新数据 → 不产生任何提交
- 支持 **Run workflow** 手动触发

注意：GitHub 定时任务可能延迟 5~20 分钟，且仓库 60 天无活动后 schedule 会暂停（手动跑一次即可恢复）。

## 本地使用

```bash
pip install openpyxl

# 增量更新数据
python3 fetch_kl8.py

# 只检查有没有新数据（有 → 退出码 0）
python3 fetch_kl8.py --check

# 本地预览 web 页面
python3 -m http.server 8000   # 打开 http://localhost:8000
```

## 数据说明

`快乐8开奖数据.json` 结构：

```json
{
  "game": "快乐8",
  "total_issues": 100,
  "date_range": { "start": "2026-07-18", "end": "2026-10-29" },
  "issues": [
    { "issue": 2026288, "date": "2026-10-29", "weekday": "周四",
      "numbers": [5, 10, 14, 19, 22, 23, 26, 27, 28, 39, 40, 47, 50, 53, 59, 61, 66, 73, 75, 79] }
  ],
  "number_frequency": { "1": 12, "2": 31 }
}
```

- `issues` 按期号倒序（最新在前）
- `numbers` 为升序的 20 个开奖号码（1~80）
- `number_frequency` 为 1~80 各号码的历史出现次数

## Web 端功能

- 最新开奖：期号、日期、20 个开奖号码
- 数据速览：总期数 / 最热号码 / 最冷号码 / 遗漏最久
- 号码频次：1~80 出现次数与频率，支持按号码 / 按热度排序，标注当前遗漏
- 历史开奖：按期号、日期、星期快速检索

## 数据来源

中国福利彩票官网开奖接口 `www.cwl.gov.cn`，仅作数据整理与展示，不构成任何投注建议。
