# 文档管理系统 (WenDang)

基于 Node.js + Express + SQLite 的部门文档管理平台，支持 AI 智能内容生成与归档。

## 功能特性

- **多部门文档管理**：支持工程部、设计部等多个部门的独立文档体系
- **AI 智能内容生成**：集成 DeepSeek API，自动将富文本资料归类到文档大纲
- **文档大纲管理**：支持两级大纲结构（大章节 + 小节），支持重置和初始化
- **富文本编辑器**：集成 WangEditor，支持图片、视频插入
- **文档解析导入**：支持 .doc, .docx, .pdf, .xls, .xlsx, .txt, .md 格式文件解析
- **版本控制**：文档快照管理，支持历史版本预览
- **权限管理**：管理员（admin）和普通编辑（editor）两种角色
- **操作日志**：记录所有用户操作行为

## 技术栈

- **后端**：Node.js + Express 5.x
- **模板引擎**：EJS
- **数据库**：SQLite3
- **文件上传**：Multer
- **文档解析**：Mammoth (Word), pdf-parse (PDF), XLSX (Excel)
- **AI 集成**：DeepSeek API

## 快速开始

### 环境要求

- Node.js >= 18
- npm

### 安装

```bash
npm install
```

### 启动

```bash
node server.js
```

服务默认运行在 `http://localhost:3000`

### 默认账号

| 用户名 | 密码 | 部门 | 角色 |
|--------|------|------|------|
| admin | admin123 | 总管理台 | 管理员 |
| gongcheng | 123456 | 工程部 | 编辑 |
| sheji | 123456 | 设计部 | 编辑 |

## 项目结构

```
wendang/
├── server.js              # 主服务文件
├── package.json           # 项目依赖配置
├── database.sqlite        # SQLite 数据库（不提交到 Git）
├── .gitignore             # Git 忽略配置
├── views/                 # EJS 模板
│   ├── index.ejs          # 前台文档展示页
│   ├── login.ejs          # 登录页
│   ├── admin.ejs          # 部门管理后台
│   └── admin_super.ejs    # 总管理台
├── public/                # 静态资源
│   ├── uploads/           # 用户上传文件
│   └── docs/              # AI 生成归档文件
└── README.md              # 项目说明
```

## 主要路由

| 路由 | 方法 | 说明 |
|------|------|------|
| `/` | GET | 首页 - 部门文档入口 |
| `/login` | GET/POST | 登录 |
| `/doc/:dept` | GET | 查看部门文档 |
| `/admin` | GET | 管理后台 |
| `/admin/init-outline` | POST | 初始化/重置大纲 |
| `/admin/generate-ai-content` | POST | AI 生成内容并归档 |
| `/admin/settings` | POST | 系统设置（API Key） |

## 配置

### AI 配置

在总管理台的"系统配置"中设置 DeepSeek API Key：

1. 使用 admin 账号登录
2. 进入"系统配置"标签
3. 填写 `deepseek_api_key`

### 大纲模板

大纲模板存储在 `settings` 表中，key 格式为 `template_{部门标识}`。

大纲格式：
- 一级标题：`一、标题内容`（使用中文数字 + 顿号）
- 二级标题：`1. 标题内容`（使用阿拉伯数字 + 点号）

## 注意事项

- `database.sqlite` 已加入 `.gitignore`，不会提交到版本控制
- `public/uploads/` 已加入 `.gitignore`，上传文件不会提交
- 首次启动会自动创建数据库和默认用户
- AI 生成内容超时时间：60 秒

## License

ISC
