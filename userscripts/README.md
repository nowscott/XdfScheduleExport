# 油猴脚本（试验版）

`xdf-schedule-export.user.js` 让你在已经登录的课表页面中，点击右下角按钮导出月视图、统计和详细课表。

## 安装与使用

1. 在浏览器扩展商店安装 Tampermonkey。
2. 点击 [一键安装油猴脚本](https://raw.githubusercontent.com/nowscott/XdfScheduleExport/main/userscripts/xdf-schedule-export.user.js)，在 Tampermonkey 安装页确认安装。
3. 正常登录 `https://we.xdf.cn/main/home` 或任意 `we.xdf.cn` 页面。
4. 点击右下角“导出课表 Excel”，使用“本周、本月、本学期”等快捷范围，或在弹窗日历中选择开始与结束日期。

脚本会从 GitHub 自动检查更新；日后不需要再次复制粘贴。

## 导出内容与恢复

默认导出顺序为：

1. `月视图`：所选范围按完整日期顺序连续展示在同一张工作表，跨月周不会拆开。每天按实际课次逐行显示，不预设时间段；取消“将多个月视图连续放在同一个工作表”后，每个月各自成为一个工作表。
2. `统计`：按学员、课程、老师、校区汇总节数。
3. `详细课表`：完整原始课程明细，可筛选、排序。

详情请求最多三路并发，以兼顾速度与接口稳定性。如果某个日期读取失败，成功部分仍会先导出，弹窗会显示“重试失败日期”按钮。脚本只在浏览器本地记住上次的日期范围和月视图选项，不会保存账号、密码或课表内容。

请只在有权访问课表的账号中使用，不要分享导出的文件或浏览器登录数据。

## 本地联调与断点调试

可以直接用 VS Code 编辑本仓库的 `xdf-schedule-export.user.js`，保存后刷新课表页面便会载入最新代码；不用反复复制粘贴。这个加载器只用于本机开发，**正式发布前请停用它，并重新启用一键安装的正式脚本**。

1. 在 Chrome/Edge 的扩展程序页打开 Tampermonkey 的“允许访问文件网址”（有些浏览器还需要把“网站访问权限”设为“所有网站”）。
2. 安装仓库中已准备好的 [本地开发加载器](./xdf-schedule-export.local-dev.user.js)：将该文件拖进浏览器，或在 Tampermonkey 的“实用工具”中导入它。它已包含本机 macOS 路径和 `xlsx` 依赖；Windows 请把最后一个 `@require` 改为 `file://C:/Users/你的用户名/.../xdf-schedule-export.user.js`。也可以手工新建脚本，正文如下：

   ```javascript
   // ==UserScript==
   // @name         XDF 课表导出（本地开发）
   // @namespace    https://github.com/nowscott/XdfScheduleCrawler
   // @match        https://we.xdf.cn/*
   // @noframes
   // @require      https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js
   // @require      file:///Users/nowscott/Documents/%E5%BC%80%E5%8F%91%E9%A1%B9%E7%9B%AE/%E5%B7%A5%E4%BD%9C%E5%B7%A5%E5%85%B7/%E7%88%AC%E5%8F%96xdf%E8%AF%BE%E8%A1%A8/userscripts/xdf-schedule-export.user.js
   // @grant        none
   // @run-at       document-idle
   // ==/UserScript==
   ```

   路径中有中文、空格或 `#` 等字符时要使用 URL 编码；在 macOS Finder 中把文件拖进浏览器地址栏后，复制得到的 `file:///...` 地址最稳妥。

3. 在 Tampermonkey 中停用正式的“XDF 课表导出”脚本，只启用“XDF 课表导出（本地开发）”。每次在 VS Code 保存后，刷新 `we.xdf.cn` 页面即可测试。

当前脚本使用 `@grant none`，因此代码运行在页面上下文中：F12 Console 保持在默认的 `top` 即可，不能也不需要切换到 Tampermonkey 上下文调用 `GM_*` API。排查时优先在源码中使用 `console.debug('[XDF 课表导出]', 变量名, 变量值)`，然后在 Console 过滤 `XDF 课表导出`。

需要暂停执行时，在要检查的源代码行单独加入 `debugger;`，先打开 F12，再刷新页面；Chrome 会在该行停住，可以查看变量、调用栈并单步执行。完成后删除 `debugger;`，再运行 `node --check userscripts/xdf-schedule-export.user.js` 和 `node --test tests/userscript.test.mjs`。
