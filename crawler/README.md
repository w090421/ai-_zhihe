# CUMT 公开信息爬虫（角色 B）

## 运行

```powershell
python -m pip install -r crawler/requirements.txt

# 全量：每个栏目近 5 页 + 全部教师
python crawler/main.py --max-pages 5 --delay 0.8

# 冒烟：1 页、每栏目 2 条、教师 3 人
python crawler/main.py --smoke --delay 0.2

# 只用缓存重算（不请求网络）
python crawler/main.py --max-pages 5 --delay 0
```

参数：

| 参数 | 默认 | 说明 |
|---|---:|---|
| `--max-pages` | 5 | 每个栏目最多列表页数 |
| `--delay` | 0.8 | 请求间隔（秒） |
| `--per-section` | 0 | 每个栏目最多条数，0 不限（调试） |
| `--teacher-limit` | 0 | 教师最多人数，0 全部（调试） |
| `--no-details` | - | 只抓列表，不抓详情 |
| `--no-teachers` | - | 跳过教师 |
| `--no-cache` | - | 禁用本地 HTML 缓存 |
| `--smoke` | - | 冒烟模式 |

## 数据源

- 学校新闻网：要闻矿大、资讯矿大、校园快讯、学术聚焦
- 学院官网：新闻动态、通知公告、学术通知
- 学院师资：4 个系的教师简介

只抓公开页面；`content.jsp` 中提示需 VPN 的页面（HTTP 403）不绕过。

## 输出

- `data/excel/矿大信息库.xlsx`
- `data/seed/school-info.csv`（团队契约：`id,title,category,summary,source_url`）
- `data/seed/school-info.full.csv`（追溯字段）
- `data/seed/teacher-info.csv`
- `data/raw/crawl_result.json`
- `data/crawl_stats.json`
- `crawler/crawl.log`

## 去重与缓存

- 按详情页 URL 去重；同一 URL 出现在多个栏目时合并栏目名。
- `data/seed/school-info.csv` 的 `id` 由 URL 哈希生成，重复运行稳定。
- 默认缓存 HTML 到 `data/cache/`（不提交 Git），中断后可继续；`--no-cache` 强制刷新。

## 站点改版

列表/详情选择器集中在 `cumt_crawler.py` 的 `parse_cs_list`、`parse_news_list`、`extract_content`、`detail_title` 等函数；改版时只改这些函数。
