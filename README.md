# 校园综合互助平台「互助地图」

把校园里的求助、组队、拼车、失物和学校信息，放到一张可检索、可接单、可追踪的地图上。

当前是 D0 完成版：页面、三层接口（类型 / 数据层 / API）、Git 流程和文档已就绪。发布、提问、失物、状态流转、归档/爬虫/汇总都能真实操作（内存演示库，重启恢复 seed）。下一步接 Supabase。

## 技术栈

- Next.js 16 + React 19 + TypeScript
- Tailwind CSS 4
- Supabase（Postgres / Auth / Storage）
- Vercel
- Leaflet（后续接真实地图）
- Python 爬虫 + 群机器人 webhook

## 目录

~~~text
.
├─ apps/web/                 Next.js 前端和 API
│  └─ src/lib/data/          数据层：全项目唯一取数入口，接 Supabase 只改这里
├─ docs/
│  ├─ 00-项目总纲.md         策划 + 分工 + 六天任务清单（正式版）
│  ├─ 01-Git协作流程.md
│  ├─ 02-Git新人指南.md
│  ├─ 03-vibe编码对接指南.md
│  ├─ api/                   接口契约
│  ├─ content/               文案和分类
│  ├─ devlog/                D0-D6 开发日志
│  └─ team/                  成员第一次 PR
├─ data/seed/                CSV/JSON 种子数据
├─ crawler/                  最小爬虫
├─ bot/                      群机器人
├─ supabase/
│  ├─ migrations/            数据库迁移
│  └─ seed/                  数据库种子
├─ .github/                  PR/Issue/CI 模板
└─ .env.example
~~~

## 本地启动

~~~powershell
cd apps/web
npm install
Copy-Item ..\..\.env.example .env.local
npm run dev
~~~

打开 http://localhost:3000。

如果还没有配 Supabase，页面会先用 mock 数据运行；填好 .env.local 后再接真实数据。

## 常用命令

~~~powershell
cd apps/web
npm run dev
npm run build
npm run lint
~~~

## 文档

- [项目总纲（正式版）](<docs/00-项目总纲.md>)
- [Git 协作流程](<docs/01-Git协作流程.md>)
- [Git 新人指南](<docs/02-Git新人指南.md>)
- [接口契约](<docs/api/接口契约.md>)
- [vibe coding 对接指南](<docs/03-vibe编码对接指南.md>)
- [分支流程（develop 集成）](<docs/05-分支流程-develop.md>)
- [开发日志 D0](<docs/devlog/D0.md>)

## 团队

| 角色 | 负责 |
| --- | --- |
| A | 架构、接口、部署、main 合并、委托后端和归档 |
| B | 学校信息库、爬虫、数据清洗、seed |
| C | 数据库、答疑、失物招领、图片、群机器人 |
| D | 委托前端、地图、筛选、响应式、共用组件 |
| E | 内容、分类、演示数据、验收、开发日志、提交材料 |

## 状态

- [x] Next.js 骨架
- [x] 页面路由和演示数据
- [x] 数据层 src/lib/data + /api/* 最小实现
- [x] vibe coding 对接指南（文件领地 + AI 提示词模板）
- [x] Git 流程和模板
- [ ] Supabase 连接
- [ ] 登录和权限
- [ ] 真实地图
- [ ] 爬虫和机器人
- [ ] P0 验收