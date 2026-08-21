# 财报跟踪平台 · 可移植版

> 本目录是「全球金属企业季度产量财报跟踪」网站的**自包含可运行版本**：代码 + 8 品种真库 Excel + 前端资源全部打包，可在**任意服务器/电脑**搭建，不依赖本地 `D:\` 路径。

## 目录结构

```
财报跟踪平台_可移植版/
├── build_site.py            # 构建脚本（Excel → data/data.js + index.html）
├── excel/                   # 8 品种真库 Excel（本目录自带，构建数据源）
│   ├── 锡.xlsx / 锌.xlsx / 铝.xlsx / 镍.xlsx
│   └── 铜.xlsx / 锂.xlsx / 硅.xlsx / 铅.xlsx
├── assets/                  # 前端资源（app.js / echarts.min.js / style.css）
├── index.html               # 站点入口
├── data/                    # 构建输入（news.json 新闻扰动 / disruptions.json 供应扰动）
├── scripts/                 # 辅助脚本（铝行业站扫描等）
├── run.py                   # 一键：构建 + 启动本地面板
├── requirements.txt         # 依赖（仅 openpyxl）
└── README_搭建说明.md
```

## 在另一台服务器搭建（3 步）

```bash
# 1. 拷贝/上传整个目录到服务器（或用 git clone 本仓库）
# 2. 装依赖
pip install -r requirements.txt        # 或 pip3 install openpyxl
# 3. 运行
python run.py                          # 构建 + 起本地面板 http://localhost:8000
python run.py --build                  # 只构建不启动（产出 data/data.js）
```

浏览器打开 `http://服务器IP:8000` 即见面板。构建输出为纯静态文件（`index.html + data/data.js + assets/`），也可整体推到任意静态托管（GitHub Pages / Nginx / 对象存储）。

## 数据源说明（Excel 唯一真库原则）

- **真库唯一在原库**（本机：`D:\拷贝文件\E\永安\{品种}\*.xlsx`），修改只改原库。
- `build_site.py` 路径解析：**真库路径存在则用真库**；不存在（如无 `D:\` 的另一服务器）则回退 `excel/` 目录内置副本。**内置副本仅供搭建演示，更新数据请改原库后重新拷贝覆盖** `excel/`。
- 更新流程：改原库 Excel → 重跑 `python build_site.py` → 重启面板即可生效。

## 与其他系统的关系

- 本机真源规范见 `D:\大胖鱼\财报汇总\caibao-collect_品种SOP_v5.md`。
- 线上外网版（本仓库）：`https://amiya866.github.io/caibao-hub/`（密码 yafco888）；旧财报平台 `https://amiya866.github.io/yafco-tracker/` 保留并行运行。
- 本可移植版与线上版同一 `build_site.py`（双模式路径），保证行为一致、单一真源。

## 常见问题

- **`openpyxl` 未安装** → `pip install -r requirements.txt`。
- **构建报错提示缺 Excel** → 确认 `excel/` 目录有 8 个文件；若在用真库路径，确认原库路径存在。
- **端口占用** → `python run.py --port 9000`。
