# PPT Template Studio

一个可在线部署的 PPT 生成工具：网页端选主题、填内容，后端用 `python-pptx` 按 JSON Schema 定义的模板实时渲染并下载 `.pptx`。

核心设计是 **结构与风格分离**：

| 概念 | 存放位置 | 描述内容 |
|---|---|---|
| **Theme（主题）** | `app/themes/*.json` | 配色、字体、字号、幻灯片尺寸 —— "长什么样" |
| **Layout（版式）** | `app/layouts/*.json` | 每种版式有哪些占位符、占位符在页面上的相对坐标、引用哪个样式令牌 —— "结构是什么" |
| **Deck spec（用户内容）** | 前端表单 / 用户传入的 JSON | 选哪个 Layout + 填什么文字 —— "内容是什么" |

同一个 Layout（如 `content_bullets`）套用 `huawei_light` 或 `apple_dark` 任意主题都能直接工作，因为 Layout 里只写 `"style": "h2"` 这种令牌，具体字号/颜色由当前主题在渲染时解析（见 `app/engine.py` 里的 `_resolve_style`）。新增一套主题不需要碰任何 Layout 或渲染代码，反之新增一种版式也不需要碰任何主题文件——这就是"可复用模板"的落地方式。

内置两套主题：
- `huawei_light` —— 华为企业标准浅色风格（红色主色调、简洁留白）
- `apple_dark` —— Apple 发布会深色风格（黑底、大字冲击、克制配色）

内置六种版式：`cover`（封面）、`toc`（目录）、`section_divider`（章节分隔）、`content_bullets`（要点列表）、`quote`（引述）、`closing`（结尾）。

## 目录结构

```
ppt-template-studio/
├── app/
│   ├── main.py              FastAPI 入口 + 路由
│   ├── engine.py             模板引擎：schema 加载 + 渲染逻辑
│   ├── themes/               主题 JSON（配色/字体/尺寸）
│   │   ├── huawei_light.json
│   │   └── apple_dark.json
│   ├── layouts/               版式 JSON（占位符坐标/样式令牌）
│   │   ├── cover.json / toc.json / section_divider.json
│   │   ├── content_bullets.json / quote.json / closing.json
│   └── static/                前端（原生 HTML/CSS/JS，无需构建）
├── examples/                  示例 deck spec（同一套版式，两套主题）
├── scripts/test_render.py      本地冒烟测试脚本
├── requirements.txt
├── Dockerfile                  容器化部署
├── render.yaml                 Render 一键部署蓝图
├── Procfile                     Railway / Heroku 风格部署
└── .gitignore
```

## 本地运行

```bash
cd ppt-template-studio
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 冒烟测试：直接渲染 examples/ 下的示例，验证模板引擎本身没问题
python scripts/test_render.py

# 启动 Web 服务
uvicorn app.main:app --reload
```

打开 `http://127.0.0.1:8000`，选择主题 → 选版式 → 填表单 → 加入大纲 → 生成并下载。右侧 JSON 面板可以直接粘贴/编辑 deck spec，用来批量导入或调试。

## 接口

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/api/themes` | 返回可选主题列表 |
| GET | `/api/layouts` | 返回可选版式及各自的占位符定义 |
| POST | `/api/generate` | 提交 deck spec，返回 `.pptx` 文件流 |
| GET | `/healthz` | 健康检查（部署平台探活用） |

`POST /api/generate` 请求体示例见 `examples/demo_deck_huawei.json`。

## 如何扩展

**新增一套主题**：在 `app/themes/` 加一个新的 `xxx.json`，字段结构照抄现有文件即可，不需要改任何代码，前端下拉框会自动出现新选项。

**新增一种版式**：在 `app/layouts/` 加一个新的 `xxx.json`，定义 `placeholders`（`key` / `type` / `style` / `rect`），`type` 目前支持 `text`、`bullet_list`、`numbered_list`；如果需要新的占位符类型（比如图片、表格、图表），在 `app/engine.py` 的 `_render_slide` 里加一个新分支即可。`style` 令牌（`title/subtitle/h2/body/caption/quote/index`）到具体字号颜色的映射在 `_resolve_style` 里维护，如果需要新的令牌，加一行就行。

## 归档到 GitHub

这份代码目前只存在于当前工作目录，需要你在自己的电脑上把它推送到 GitHub（这一步涉及你的 GitHub 账号授权，无法代劳）：

```bash
cd ppt-template-studio
git init
git add .
git commit -m "Init: PPT template studio"

# 先在 github.com 新建一个空仓库（不要勾选自动生成 README），拿到仓库地址后：
git branch -M main
git remote add origin https://github.com/<你的用户名>/ppt-template-studio.git
git push -u origin main
```

## 部署上线

**重要澄清**：GitHub 本身（GitHub Pages）只能托管纯静态网站，不能运行这个项目的 FastAPI 后端。业内通常说的"部署在 GitHub 上"，实际做法是把代码推到 GitHub 仓库，再用一个支持后端运行时的平台"连接"这个仓库，由它们从 GitHub 自动拉取代码构建部署。项目里已经准备好了两种平台的配置文件，任选其一：

### 方式一：Render（推荐，配置最简单）

1. 访问 [render.com](https://render.com)，用 GitHub 账号登录并授权访问你的仓库。
2. New → Web Service → 选择 `ppt-template-studio` 仓库。
3. Render 会自动识别根目录的 `render.yaml`（Blueprint），直接按 Dockerfile 构建，无需手填 Build/Start 命令。
4. 部署完成后会给你一个 `https://xxx.onrender.com` 的公网地址。
5. 之后每次 `git push` 到 `main` 分支，Render 会自动重新构建部署。

### 方式二：Railway

1. 访问 [railway.app](https://railway.app)，New Project → Deploy from GitHub repo → 选择该仓库。
2. Railway 会识别 `Procfile`（或直接用 Dockerfile），自动安装依赖并启动。
3. 在 Settings → Networking 里生成一个公网域名。

两种方式都是"连接 GitHub 仓库即自动部署"，代码改动、`git push` 之后平台会自动重新构建，不需要手动上传文件。

## 已知限制 / 后续可扩展方向

- 当前版式为文字类布局，暂未支持图片/图表占位符（`engine.py` 的类型分支设计已预留扩展空间）。
- 前端大纲编辑是表单式，不是所见即所得的可视化拖拽编辑器（如需要，可在此基础上叠加一个基于坐标的画布编辑器）。
- 生成的 `.pptx` 是全新文件，不支持"编辑已有 pptx 的部分内容"。
