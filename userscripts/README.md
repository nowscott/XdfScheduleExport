# 油猴脚本（试验版）

`xdf-schedule-export.user.js` 让你在已经登录的课表页面中，点击右下角按钮导出指定日期范围的课程明细与五时段月视图。

## 安装与使用

1. 在浏览器扩展商店安装 Tampermonkey。
2. 点击 [一键安装油猴脚本](https://raw.githubusercontent.com/nowscott/XdfScheduleCrawler/main/userscripts/xdf-schedule-export.user.js)，在 Tampermonkey 安装页确认安装。
3. 正常登录 `https://we.xdf.cn/main/home` 或任意 `we.xdf.cn` 页面。
4. 点击右下角“导出课表 Excel”，在弹窗日历中选择开始与结束日期。

脚本会从 GitHub 自动检查更新；日后不需要再次复制粘贴。

## 与 Python 版的差异

油猴版会生成“课表”明细表，以及每个月一张固定五时段的月视图。它不会保存账号或密码；请求直接使用你当前浏览器已登录页面的会话。

请只在有权访问课表的账号中使用，不要分享导出的文件或浏览器登录数据。
