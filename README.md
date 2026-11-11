
# 📑 Word Document Tabs

**A lightweight, free and customizable vertical document tab manager for Microsoft Word on Windows.**

一款轻量、免费、可自定义的 Microsoft Word 竖向文档标签管理工具。

![Windows](https://img.shields.io/badge/Platform-Windows-blue)
![AutoHotkey](https://img.shields.io/badge/AutoHotkey-v2-green)
![Status](https://img.shields.io/badge/Status-Beta-orange)

[English](#english) | [简体中文](#简体中文)

---

## 简体中文

### ✨ 项目介绍

Word Document Tabs 是一款基于 AutoHotkey v2 开发的 Windows 桌面辅助工具。

它为 Microsoft Word 提供独立的悬浮式文档标签栏，让用户可以更方便地切换和管理多个 Word 文档窗口。

无需修改 Word 安装文件，也无需购买付费标签页插件。

### 🚀 主要功能

- 📑 自动识别已打开的 Word 文档窗口
- 🖱️ 点击标签快速切换文档
- 🎨 鼠标悬停时显示动态高亮效果
- 📌 当前文档保持高亮显示
- ↕️ 支持拖动标签调整排列顺序
- 💾 保存自定义标签排列顺序
- 🪟 支持悬停显示文档缩略预览
- 📐 标签栏可以自由移动和调整大小
- 📍 支持重新贴靠 Word 窗口右侧
- 💻 适用于 Windows 桌面环境

> 注意：部分动画、拖动排序及预览功能仍处于测试阶段。

### 📥 安装方法

1. 下载并安装 [AutoHotkey v2](https://www.autohotkey.com/)。
2. 从本仓库下载 `Word_Document_Tabs_Animated_Order.ahk`。
3. 双击运行该脚本。
4. 打开两个或更多 Microsoft Word 文档。
5. 使用标签栏切换和管理文档。

### ⌨️ 快捷键

| 快捷键 | 功能 |
|---|---|
| Ctrl + Alt + W | 显示或隐藏标签栏 |
| Ctrl + Alt + D | 重新贴靠到 Word 右侧 |

### 🖥️ 系统要求

- Windows 10 / Windows 11
- AutoHotkey v2
- Microsoft Word 桌面版

本项目主要针对 Word 2019、2021、2024 开发。Microsoft 365 桌面版的兼容性仍需进一步验证。

不支持 macOS 版 Word 或 Word 网页版。

### ⚠️ 已知限制

- 标签栏是独立的悬浮窗口，并非嵌入 Word 内部的原生插件。
- 标签栏可能覆盖部分 Word 界面。
- 部分 Word 窗口可能无法显示实时缩略图。
- 不同屏幕缩放比例可能影响窗口定位。
- 动画和拖动排序功能仍需进一步测试。

---

## English

### Overview

Word Document Tabs is a lightweight Windows utility built with AutoHotkey v2.

It provides a floating vertical tab panel for managing multiple Microsoft Word document windows.

The application works alongside Microsoft Word without modifying its installation files.

### Features

- Detect open Microsoft Word windows
- Switch between documents with one click
- Animated hover highlighting
- Persistent active-document highlighting
- Drag-and-drop tab reordering
- Save custom document order
- Document thumbnail previews on hover
- Movable and resizable floating panel
- Right-side docking shortcut
- Lightweight Windows utility

Some features are experimental and require additional testing.

### Installation

1. Install [AutoHotkey v2](https://www.autohotkey.com/).
2. Download `Word_Document_Tabs_Animated_Order.ahk`.
3. Double-click the script to run it.
4. Open multiple Microsoft Word documents.
5. Use the floating tab panel to switch between them.

### Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| Ctrl + Alt + W | Show / Hide panel |
| Ctrl + Alt + D | Re-dock panel |

### Compatibility

Designed for Windows 10 and Windows 11 with desktop Microsoft Word.

Microsoft 365 compatibility has not yet been fully tested.

### Limitations

This project is an independent floating window manager, not a native Microsoft Word add-in.

Some document previews, animations and drag-and-drop behaviors may vary depending on the system environment.

---

## 🛠️ Built With

- AutoHotkey v2
- Windows API
- Windows Desktop Window Manager (DWM)

## 📌 Project Status

**Beta — Under active development**

Bug reports, feedback and suggestions are welcome through GitHub Issues.

## 👨‍💻 Author

Created by [LiXinze-FEM](https://github.com/LiXinze-FEM).

## Disclaimer

This is an independent open-source project and is not affiliated with, endorsed by, or sponsored by Microsoft.

Microsoft Word is a trademark of Microsoft Corporation.
