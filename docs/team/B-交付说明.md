# 校园综合互助平台 · 角色 B 交付（学校信息库 / 爬虫）

工作目录：`sites/ai_zhihe`。本目录对应角色 B（王禹涵）的领地：`crawler/`、`data/seed/school-info.csv`。

## 1. 本次交付物

| 文件 | 用途 |
|---|---|
| `data/excel/矿大信息库.xlsx` | 面向本次任务的 Excel 信息库，含学校新闻、学院动态、教师信息、数据说明 4 个 Sheet |
| `data/seed/school-info.csv` | 团队契约格式，供前端/Supabase seed 使用 |
| `data/seed/school-info.full.csv` | 追溯版，保留原始栏目、来源、发布日期、标签 |
| `data/seed/teacher-info.csv` | 教师简表（姓名、院系、职称、导师类别、研究方向等） |
| `crawler/cumt_crawler.py` | 爬虫实现（列表分页 + 详情 + 教师 + 去重 + Excel/CSV 输出） |
| `crawler/main.py` | 兼容入口，命令与团队 README 对齐 |
| `crawler/crawl.log` | 最近一次运行日志 |
| `data/crawl_stats.json` | 统计、分栏目数量、失败请求 |
| `data/raw/crawl_result.json` | 本次抓取的完整结构化结果（可复现） |
| `docs/数据来源与合规说明.md` | D5 合规说明 1 页 |

## 2. 数据规模（本次运行）

- 新闻/通知：**310 条**（学校 180，学院 130）
- 教师：**127 人**（计算机科学与技术系 49、网络空间安全系 23、数据科学与工程系 19、人工智能系/智能科学与技术系 36）
- 列表页：7 个栏目 × 近 5 页
- 详情页：310 篇新闻/通知 + 127 位教师
- 失败请求：0
- 发布日期范围：2023-12-12 ~ 2026-11-26
- 去重：按详情页 URL；本次 310 条 URL 全部唯一

分栏目：

| 站点 | 栏目 | 条数 |
|---|---|---:|
| 学校 | 要闻矿大 | 45 |
| 学校 | 资讯矿大 | 45 |
| 学校 | 校园快讯 | 45 |
| 学校 | 学术聚焦 | 45 |
| 学院 | 新闻动态 | 30 |
| 学院 | 通知公告 | 50 |
| 学院 | 学术通知 | 50 |

## 3. 快速运行

```powershell
python -m pip install -r crawler/requirements.txt

# 全量（每栏目近 5 页 + 全部教师）
python crawler/main.py --max-pages 5 --delay 0.8

# 冒烟测试：1 页、每栏目 2 条、教师 3 人
python crawler/main.py --smoke --delay 0.2

# 只抓列表，不抓详情
python crawler/main.py --max-pages 3 --no-details
```

本爬虫默认开启本地 HTML 缓存（`data/cache/`），中断后重跑不会重复请求已完成页面；需要强制刷新时加 `--no-cache`。

## 4. Excel 结构

`data/excel/矿大信息库.xlsx`：

- **学校新闻**：序号、栏目、分类、标题、发布日期、摘要、正文、作者、来源、链接、抓取时间
- **学院动态**：序号、栏目、分类、标题、发布日期、摘要、正文、来源、链接、抓取时间
- **教师信息**：序号、姓名、所属院系、职称、导师类别、研究方向、毕业院校、邮箱、简介、发布日期、详情链接、来源、抓取时间
- **数据说明**：数据源、抓取范围、去重规则、合规说明、字段说明、维护方式

链接列已设置为可点击超链接；表头冻结并带筛选。

## 5. 团队契约（school-info.csv）

与仓库现有 `data/seed/school-info.csv` 表头保持一致：

```csv
id,title,category,summary,source_url
```

- 编码：UTF-8，**无 BOM**（与仓库现有 seed 一致）
- `id`：`cumt-news-<URL sha1 前 10 位>`，重复跑保持稳定
- `category`：已映射到 `InfoCategory` 枚举：`教务 / 生活 / 场馆 / 奖助 / 就业 / 其他`
  - 通知公告、学术通知默认 `教务`
  - 标题命中招聘/就业/双选等 → `就业`；奖助 → `奖助`；场馆 → `场馆`；生活 → `生活`
  - 其余新闻 → `其他`
- 去重：按 `source_url`；重复跑全量刷新，不累加
- 条数：310 条，满足 P0「信息库 ≥ 50 条，每条有来源」

`school-info.full.csv` 额外保留：原始分类、映射后分类、来源名称、发布日期、栏目标签。

## 6. 与团队仓库对接

角色 B 只需把以下内容放入团队仓库根目录：

```text
crawler/
  cumt_crawler.py
  main.py
  requirements.txt
  run_crawler.ps1
  schedule-github-actions.yml   # 示例，实际放到 .github/workflows/
data/seed/
  school-info.csv               # 必需，契约格式
  school-info.full.csv          # 可选，追溯用
  teacher-info.csv              # 可选，教师简表
docs/数据来源与合规说明.md
```

前端 `apps/web/src/lib/data/info.ts` 当前是 mock，接入时把 `school-info.csv` 导入 Supabase 或读取为 `SchoolInfo[]` 即可；字段映射：

```text
SchoolInfo.id        <- id
SchoolInfo.title     <- title
SchoolInfo.category  <- category
SchoolInfo.summary   <- summary
SchoolInfo.source    <- source_url
SchoolInfo.updatedAt <- (导入时间，或 published_at)
```

`/api/internal/crawl` 目前是演示桩，可在管理页触发 Python 爬虫（本地/CI/定时任务均可）。

## 7. 定时任务

Windows 任务计划程序：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\path\to\crawler\run_crawler.ps1"
```

GitHub Actions 示例见 `crawler/schedule-github-actions.yml`，建议每天一次（例如 UTC 20:00 / 北京时间 04:00）刷新并提交 `data/seed`。

## 8. 已知边界

- `content.jsp?...` 的 5 条通知返回「请登录客户端 VPN 后浏览」（HTTP 403），**未绕过**，仅保留标题和链接，正文为空。
- 13 位教师的详情页显示「正在建设...」，简介为空。
- 19 位教师（含英文简介、页面未写研究方向）未提取到研究方向。
- 教师邮箱只在官网公开时采集；未采集手机号、住址等个人敏感信息。
- 站点改版可能导致选择器失效，失败时保留 seed 兜底。

合规细节见 [`docs/数据来源与合规说明.md`](docs/数据来源与合规说明.md)。
