# 爬取xdf课表

GitHub 仓库名建议：`XdfScheduleCrawler`

一个本地运行的 XDF 课表抓取与 Excel 导出工具。程序通过 Playwright 打开课表页面，复用用户登录状态调用课表接口，并生成便于查看和筛选的工作簿。

## 功能

- 在浏览器中手动登录，不在代码中保存账号或密码
- 按任意日期范围批量抓取课程
- 跨月日期范围会自动按月份拆分请求
- 自动跳过无课日期，并按日期、时间排序
- 导出完整课表明细
- 为每个月生成日历式月视图
- 每个日期固定显示 5 个时段，一节课占一行
- 绿色表示有课，空闲时段明确标注“无课”
- 浏览器登录状态保存在本机，方便下次继续使用

## 环境要求

- macOS、Windows 或 Linux
- Python 3.10+
- 可访问 XDF 课表页面的账号

## 安装

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m playwright install chromium
```

Windows PowerShell 激活虚拟环境：

```powershell
.venv\Scripts\Activate.ps1
```

## 使用

交互输入日期：

```bash
python main.py
```

直接指定日期：

```bash
python main.py --start 2026-07-13 --end 2026-08-31
```

macOS/Linux 也可以使用启动脚本：

```bash
./run.sh --start 2026-07-13 --end 2026-08-31
```

首次运行会打开浏览器，请完成登录。登录状态有效时，后续运行一般不需要再次登录。

## 输出

文件默认保存在 `output/`，每个 Excel 包含：

- `课表`：完整课程明细，可筛选、排序
- `YYYY年M月月视图`：星期一到星期日的月历布局

月视图中每个日期固定包含以下 5 行：

1. 08:00 时段
2. 10:00/10:20 时段
3. 13:40 时段
4. 16:00 时段
5. 18:00/18:30 时段

实际有课时显示真实的起止时间、学生和教室编号；没有课程时显示“无课”。

## 可选环境变量

| 变量 | 用途 |
| --- | --- |
| `XDF_USER_DATA_DIR` | 浏览器登录数据目录 |
| `XDF_EXPORT_DIR` | Excel 输出目录 |
| `XDF_LOGIN_TIMEOUT` | 等待登录秒数，默认 180 |
| `XDF_REQUEST_INTERVAL` | 每次接口请求的间隔秒数，默认 0.4 |
| `XDF_SCHEDULE_PAGE_URL` | 覆盖课表页面地址 |
| `XDF_API_BASE` | 覆盖接口基础地址 |

## 隐私与安全

`browser_data/` 中可能包含登录 Cookie 和其他浏览器会话数据，已经加入 `.gitignore`。不要把这个目录、真实课程 Excel、接口响应 JSON 或截图提交到 GitHub。

本项目仅用于整理本人有权访问的课表数据。请遵守所在机构的使用规则，不要高频请求、共享他人信息或将工具用于未授权的数据获取。

## 项目结构

```text
XdfScheduleCrawler/
├── main.py             # 命令行入口
├── crawler.py          # 浏览器登录与课表抓取
├── excel_exporter.py   # 明细表及五时段月视图
├── config.py           # 路径、接口和运行参数
├── requirements.txt    # Python 依赖
├── run.sh              # macOS/Linux 启动脚本
└── tests/              # 本地单元测试
```
