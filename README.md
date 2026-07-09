# NIUERA 报价工作台 · Quote Workbench

一个本地运行的报价制作应用:浏览器界面里选产品、填客户信息(关键字段已预填默认值)、
一键生成双语报价 xlsx。两种模式:**标准**(数量+合计)、**分项**(按量分档)。

---

## 一、最快上手(开发/试用,无需打包）

需要本机有 Python 3.10+。

```bash
# 1. 进入项目目录,装依赖(只需一次)
pip install -r requirements.txt

# 2. 启动
#    Windows: 双击 start.bat
#    Mac/Linux: ./start.sh
#    或手动:
python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

浏览器打开 http://127.0.0.1:8000 即是工作台。

**日常报价 3 步:** 填客户公司 → 左侧目录点「添加」选产品、填数量 →「生成报价 (xlsx)」。
PI 编号、日期、From、商务条款都已自动预填,改不改都行。

---

## 二、打包成单个 .exe（给非技术同事 / 自己免环境）

> exe 必须在 **Windows** 上打包(PyInstaller 不能跨平台产出 exe)。

```bat
:: 在 Windows 上,项目目录里双击或运行:
build_exe.bat
```

完成后 exe 在 `dist\NIUERA报价工作台.exe`,双击即自动起服务并打开浏览器。

### 想让 exe 用户随时改产品库?
默认 `build_exe.bat` 把 `data\` 打进了 exe(改产品要重新打包)。
若希望 `products.json` / `config.json` 放在 exe 旁边、可随时编辑,改用「外置数据」:
1. 打包时去掉 `--add-data "data;data"`;
2. 把 `data\` 文件夹和 exe 放同一目录发出去;
3. `app.py` 已按「exe 同目录」找 `data\`,无需改代码。

---

## 三、目录结构

```
niuera-quote-app/
├─ app.py              FastAPI 后端(路由、生成、产品管理)
├─ launcher.py         exe 入口(起服务+开浏览器)
├─ engine/
│  └─ quote_engine.py  报价生成引擎(openpyxl,双语排版,动态行数)
├─ static/
│  └─ index.html       网页界面(原生 JS,零依赖)
├─ data/
│  ├─ products.json    产品库(56 个型号,可在界面「产品管理」里增删改)
│  ├─ config.json      公司抬头 + 默认商务条款 + PI 自增号
│  ├─ logo.jpeg        Logo
│  └─ backups/         每次保存自动备份
├─ output/             生成的报价文件
├─ requirements.txt
├─ start.bat / start.sh
└─ build_exe.bat
```

---

## 四、要点

- **预填默认值**:From、币种、Origin、Payment、Price Term、Delivery Time、Package、Validity
  都来自 `config.json` 的 `defaults`;PI 号 = 前缀 + 自增序号,生成后自动 +1。改默认值改这里。
- **两套产品名**:选型名(红,带型号便于区分,**不进报价单**)/ 对外名(报价单 Item 列显示)。
- **动态行数**:几个产品就几行,没有多余空行(根治了 Excel 模板里删空行的麻烦)。
- **PDF 导出**:界面里「导出 PDF」需本机装 LibreOffice;没装就导 xlsx 后在 Excel 里另存为 PDF。
- **产品管理**:界面右上「⚙ 产品管理」可增删改产品、调价,保存写回 `products.json`(自动备份)。
- **接 n8n**:`engine/quote_engine.py` 是独立模块,`POST /api/generate` 是标准接口,
  后续可让 n8n 在收到询盘后直接调它自动出报价。

---

## 五、字段格式(产品库)

| 字段 | 说明 |
|---|---|
| `select` | 选型名(唯一,下拉/搜索用,不进报价) |
| `item` | 对外名(报价单显示) |
| `category` | 类别 |
| `desc` | 英文描述(支持换行) |
| `unit` | 单位(PCS) |
| `band1`/`price1` | 数量档①与单价①(分项模式用;标准模式默认单价取 price1) |
| `band2`/`price2` | 数量档②与单价②(只有一档就留空) |
| `model` | 内部型号(仅记录,不显示) |

---

## V4:报价 + PI 双单据系统(2026-07)

在原报价工作台基础上升级为双单据系统:

- **PI(Proforma Invoice)生成**:七列表头、费用行跨列合并、英文大写金额(服务端权威计算 + 前端高亮人核)、红色 Validity、银行信息块、可选透明电子章。
- **报价一键转 PI**:历史报价的客户/明细/价格全继承,可沿用编号(同号一份报价 + 一份 PI)或生成时取新号;阶梯报价转单时强制核对数量与成交价。
- **单据存档与历史检索**:所有生成的单据自动入档(`data/documents/`,每单一个 JSON + 可重建索引),按类型/报价人/客户/日期筛选,支持重新下载。
- **只读快照 + 版本机制**:单据生成后不可改,修改走「复制为新版」生成 `-R1/-R2`,原档永不覆盖。
- **并发安全**:跨进程文件锁(Windows/Linux 兼容,含陈旧锁自愈),生成时锁内取号,多人同时开单不撞号。
- **多报价人**:报价人下拉 + 浏览器本机记忆,历史页按人分类,无需账号体系。

### 目录结构

```
app.py                 FastAPI 后端(唯一写路径 POST /api/generate)
engine/quote_engine.py 报价渲染
engine/pi_engine.py    PI 渲染
engine/docstore.py     存档/索引/文件锁
engine/amount_words.py 金额转英文大写(tests/ 含边界用例)
tools/                 config 升级、历史回填、路径迁移脚本
static/index.html      单页前端(报价 | PI | 历史 | 设置)
deploy/                Docker + Caddy 云端部署模板(见 Caddyfile.example)
```

### 云端部署(可选)

`deploy/` 提供 Docker Compose 模板:应用容器(含 LibreOffice,支持服务端直出 PDF)+ Caddy 自签 HTTPS 反代。国内服务器构建时 Dockerfile 已内置腾讯云内网 apt/pip 源。首次使用:复制 `deploy/Caddyfile.example` 为 `Caddyfile` 并填入服务器 IP,`docker compose up -d --build`。

> 声明:本仓库不含任何真实业务数据。`data/` 下的银行账户、产品价格、客户档案、电子章均被 .gitignore 排除,请按 `data/README.md` 用 example 模板初始化。
