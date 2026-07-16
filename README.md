# XDF 课表导出工具

[![Tampermonkey](https://img.shields.io/badge/Tampermonkey-Userscript-00485B?logo=tampermonkey&logoColor=white)](https://www.tampermonkey.net/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

一个浏览器本地运行的油猴脚本。它直接复用你在课表页面中的登录状态，生成便于查看和筛选的 Excel 工作簿；无需安装 Python 或运行本地服务。

> 非官方项目，与新东方及其关联方无关。本项目不提供账号、绕过登录的功能，也不应被用于访问无权查看的数据。

## 功能

- 一键安装后，在已登录的 `we.xdf.cn` 页面直接使用
- 日期快捷选择：本周、下周、本月、下月、本学期
- 记住上一次导出的日期范围和月视图偏好（仅浏览器本地）
- 跨月日期范围会自动按月份拆分请求，课程详情最多三路并发读取
- 某些日期读取失败时先导出成功部分，并可一键重试失败日期
- 导出顺序为：月视图、统计、详细课表
- 默认按日期连续生成跨月月视图：跨月周不会拆开，也可关闭该选项拆分为每月一张
- 月视图按每周真实时间段排布：同一横排只会显示相同的起止时间，不预设或限制时间段
- 统计页按学员、课程、老师、校区汇总节数

## 安装

1. 在浏览器扩展商店安装 Tampermonkey。
2. 点击 [一键安装油猴脚本](https://raw.githubusercontent.com/nowscott/XdfScheduleExport/main/userscripts/xdf-schedule-export.user.js)。
3. 正常登录 `https://we.xdf.cn/main/home` 或任意 `we.xdf.cn` 页面。
4. 点击右下角“导出课表 Excel”，选择日期范围或使用快捷范围。

脚本会通过 GitHub 自动检查更新，日后不需要再次复制粘贴。

## 输出

每个 Excel 依次包含：

- `月视图`：默认按日期顺序连续排布在一张工作表，跨月周保持在同一行，顶部仅显示一次完整导出范围；关闭“同一工作表”选项后则为每月一张。
- `统计`：按学员、课程、老师、校区汇总节数，并显示范围、课时、时长等概览。
- `详细课表`：完整原始课程明细，可筛选、排序。

月视图会根据每周实际出现的时间段动态增加课程行，最左侧固定显示每行的时间段，并显示真实的学生和教室编号；同一横排只对应同一个起止时间，缺少该时段的日期会留空。没有课程的日期仅显示一行“无课”。

## 隐私与安全

脚本不会保存账号、密码或课程内容。它只会在浏览器本地保存你上一次使用的日期范围与月视图布局选项。不要把真实课程 Excel、接口响应 JSON 或截图提交到 GitHub。

本项目仅用于整理你有权访问的课表数据。请遵守所在机构的使用规则，不要高频请求、共享他人信息，或将工具用于未授权的数据获取。

## 贡献

欢迎提交 Issue 和 Pull Request。提交前请阅读 [贡献指南](CONTRIBUTING.md)，尤其注意不要上传登录态、真实课表或任何个人信息。

## 许可证

本项目采用 [MIT License](LICENSE) 开源。

## 更新日志

详见 [CHANGELOG.md](CHANGELOG.md)。

## 项目结构

```text
XdfScheduleExport/
├── userscripts/
│   ├── xdf-schedule-export.user.js  # 油猴脚本
│   └── README.md                     # 安装与使用说明
├── CHANGELOG.md
└── LICENSE
```
