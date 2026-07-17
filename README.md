# 林的记事本

一个本地桌面记事本/任务记录工具。

## 当前功能

- 新增、编辑、删除任务
- 设置计划日期
- 手动点击开始，记录开始时间
- 手动点击结束，记录结束时间，并标记完成
- 自动区分今天、三天内提醒、未来、已完成
- 写随笔，记录开始写作时间
- 随笔窗口支持一键复制提示词并打开 ChatGPT 网页
- 随笔窗口支持粘贴 ChatGPT 回应，并自动清理空行、Markdown 符号和多余换行
- 任务完成记录和随笔记录会合并到同一个 Word：`life_journal.docx`
- 本地 SQLite 保存任务和随笔数据
- 主窗口使用半透明高对比主题，适合常驻桌面
- `设置` 下拉菜单可调整字体颜色、字体透明度和背景透明度
- 支持窗口置顶
- 支持 Windows 开机自启动开关

## 运行

在 PowerShell 里运行：

```powershell
cd path\to\daily_notebook
.\run.ps1
```

第一次运行会创建 `.venv` 并安装依赖。

## 文件

- `app.py`: 主程序
- `tasks.db`: 运行后自动生成的本地数据库，不会上传到 GitHub
- `life_journal.docx`: 统一的 Word 生活记录，包含已完成任务和随笔，不会上传到 GitHub
- `图片1.png`: 当前图标源图
- `daily_notebook.ico`: 桌面快捷方式图标
- `make_icon_and_shortcut.py`: 重新生成图标和桌面快捷方式 `林的记事本.lnk`
- `start_hidden.vbs`: 隐藏 PowerShell 窗口启动程序
- `run.ps1`: 启动脚本
- `requirements.txt`: Python 依赖

## ChatGPT 辅助回应

在随笔窗口里点击 `复制并打开 ChatGPT`，程序会把当前随笔整理成一段提示词复制到剪贴板，并打开 `https://chatgpt.com/`。

进入网页后直接粘贴发送即可。这样可以使用你平时的 ChatGPT 对话环境和它已有的上下文。

拿到 ChatGPT 的回答后，回到随笔窗口点击 `粘贴 ChatGPT 回应`，把回答粘贴进去。程序会自动：

- 去掉空白行
- 去掉表情和代码块
- 保留 ChatGPT 回答的段落、重点块和列表结构
- 遇到句号、问号、感叹号时适度重新分行
- 把清理后的内容写入 `life_journal.docx`
