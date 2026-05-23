# Personal Finance Tracker

> 个人记账系统 — Python + PyQt5 桌面应用

## 项目类型

Python PyQt5 桌面 GUI 应用程序，纯本地运行，数据存储在 Excel 文件中。

## 技术栈

| 技术 | 用途 |
|------|------|
| Python | 编程语言 |
| PyQt5 | 图形界面框架 |
| openpyxl | Excel 文件读写（数据存储） |
| matplotlib | 统计图表绘制 |

## 推荐 Python 版本

**Python 3.10 ~ 3.12**

> Python 3.12 经验证可正常运行。Python 3.14 不兼容 PyQt5，请勿使用。

## 安装与运行

### 1. 创建虚拟环境

```bash
# Windows（使用 Python 3.10、3.11 或 3.12）
py -3.12 -m venv .venv
# 或
py -3.11 -m venv .venv
```

### 2. 激活虚拟环境

```bash
# Windows
.venv\Scripts\activate
```

### 3. 安装依赖

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 4. 启动项目

```bash
python main.py
```

> main.py 已内置 Qt 平台插件路径自动检测，不需要再手动设置环境变量。

## 默认账号

| 用户名 | 密码 |
|--------|------|
| `0` | `1111` |
| `User1` | `1111` |

> 注：readme.txt 中提到的 User1 与 Users.xlsx 中实际存储的用户名 `0` 不同，以 Users.xlsx 为准。

## 功能列表

- 用户登录 / 注册
- 添加账单（金额、分类、日期、备注）
- 导入账单（Excel 文件）
- 查看账单明细（表格形式）
- 查询账单（按日期、分类筛选）
- 修改账单（表格内直接编辑）
- 删除账单（勾选批量删除）
- 撤回（撤销最后一条记录）
- 切换账户
- 统计：今日 / 本周 / 本月 / 本年 / 总支出（含柱状图）

## 常见问题

### PyQt5 与 Python 3.14 不兼容

PyQt5 的 C 扩展不支持 Python 3.14。如果系统默认 Python 为 3.14，请使用 `py -3.10` 或 `py -3.11` 指定版本创建虚拟环境。

### 启动时报 Segmentation fault

通常是因为 PyQt5 与当前 Python 版本不兼容。请确认使用的是 Python 3.8 ~ 3.11。

### 找不到 xlsx 文件

请确保在项目根目录下运行 `python main.py`，不要从其他目录启动。程序已使用 `BASE_DIR` 自动解析文件路径，但仍建议在项目根目录启动。

### Qt 平台插件错误（qt.qpa.plugin: Could not find the Qt platform plugin）

此错误已在 main.py 中通过自动设置 `QT_QPA_PLATFORM_PLUGIN_PATH` 解决。如果仍出现此问题：

1. 确认使用的是项目 .venv 中的 Python（`.venv\Scripts\python.exe`），不是系统 Python 3.14
2. 如果 .venv 不存在或损坏，重新创建：`py -3.12 -m venv .venv`
3. 然后 `pip install -r requirements.txt`

## 项目结构

```
├── main.py               # 项目入口，全部业务逻辑
├── Login.py              # 登录界面（由 Login.ui 自动生成）
├── BillingSystem.py      # 主界面（由 BillingSystem.ui 自动生成）
├── Bill.py               # 账单明细界面（由 Bill.ui 自动生成）
├── Statistic.py          # 统计界面（由 Statistic.ui 自动生成）
├── requirements.txt      # Python 依赖
├── Users.xlsx            # 用户账号数据
├── User1.xlsx            # 演示账单数据
├── ImportData.xlsx       # 导入功能演示文件
├── ui/                   # Qt Designer 界面文件
│   ├── Login.ui
│   ├── BillingSystem.ui
│   ├── Bill.ui
│   └── Statistic.ui
└── readme.txt            # 原始项目说明
```
