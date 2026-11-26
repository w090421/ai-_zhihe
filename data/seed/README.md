# seed 数据（角色 B）

由 `crawler/cumt_crawler.py` 生成，请勿手工编辑；重新运行爬虫即可全量刷新。

| 文件 | 编码 | 说明 |
|---|---|---|
| `school-info.csv` | UTF-8 无 BOM | **团队契约格式**：`id,title,category,summary,source_url` |
| `school-info.full.csv` | UTF-8 with BOM | 追溯版：保留原始分类、映射分类、来源名称、发布日期、栏目 |
| `teacher-info.csv` | UTF-8 with BOM | 教师简表：姓名、院系、职称、导师类别、研究方向、毕业院校、邮箱、简介、链接 |

## school-info.csv 字段

| 字段 | 说明 |
|---|---|
| `id` | `cumt-news-<URL sha1 前 10 位>`，稳定可重复 |
| `title` | 标题 |
| `category` | `教务 / 生活 / 场馆 / 奖助 / 就业 / 其他` |
| `summary` | 列表摘要或正文首个实质段落前 180 字 |
| `source_url` | 官网原文链接，按此去重 |

## 重新生成

```powershell
python crawler/main.py --max-pages 5 --delay 0.8
```

本次数据规模：310 条新闻/通知、127 位教师；来源与合规说明见 `docs/数据来源与合规说明.md`。
