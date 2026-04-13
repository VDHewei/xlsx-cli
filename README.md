# xlsx-cli

**轻量级 Office 文件编辑器 | Lightweight Office File Editor**

基于 [Bun](https://bun.sh/) 运行时的 XLSX / DOCX 读写工具，支持 **CLI 命令行** 和 **Webview GUI 图形界面** 双模式。

A lightweight XLSX / DOCX editor powered by [Bun](https://bun.sh/), supporting both **CLI** and **Webview GUI** modes.

[English](#english) | [中文](#-简介)

---

## 简介

xlsx-cli 是一款基于 Bun 构建的跨平台 Office 文件编辑器，提供以下核心能力：

- **XLSX 读写引擎** — 基于 SheetJS，完整保留单元格样式（字体、颜色、对齐、合并等）
- **DOCX 读写引擎** — 基于 `docx` 库 + 自研 ZIP/XML 解析器，支持段落样式与图片提取
- **Webview GUI** — 原生桌面窗口，内嵌 SPA 前端，零外部前端依赖
- **CLI 模式** — 适合自动化批处理场景
- **国际化** — 支持简体中文、繁體中文、English、日本語 四种语言
- **跨平台编译** — 支持 Windows (.exe)、macOS (.app)、Linux 二进制打包

## Features

- **XLSX Engine** — Powered by SheetJS with full cell style preservation (font, color, alignment, merge, etc.)
- **DOCX Engine** — Built on `docx` library with custom ZIP/XML parser, supports paragraph styling & image extraction
- **Webview GUI** — Native desktop window with embedded SPA, zero external frontend dependencies
- **CLI Mode** — Ideal for automation and batch processing
- **i18n** — Supports Simplified Chinese, Traditional Chinese, English, and Japanese
- **Cross-platform Build** — Windows (.exe), macOS (.app), Linux binary packaging

## 快速开始 | Quick Start

### 环境要求 | Prerequisites

- [Bun](https://bun.sh/docs/installation) >= 1.x

### 安装 | Installation

```bash
git clone https://github.com/VDHewei/bun-cli.git
cd bun-cli
bun install
```

### 运行 | Run

```bash
# GUI 模式（默认）
bun run start

# CLI 模式
bun run src/index.ts xlsx <file.xlsx>
bun run src/index.ts docx <file.docx>
```

## 使用方法 | Usage

### GUI 模式 | GUI Mode

启动后自动打开原生桌面窗口，支持：
- **拖拽上传** XLSX / DOCX 文件
- **电子表格编辑** — 实时编辑单元格、合并单元格、添加行列、应用规则
- **文档编辑** — 编辑段落内容与样式、添加/删除段落
- **设置页面** — 语言切换、自定义规则配置

### CLI 模式 | CLI Mode

```bash
# 读取并显示 XLSX 内容
bun run src/index.ts xlsx path/to/file.xlsx

# 显示文件元信息（工作表数、行列数等）
bun run src/index.ts xlsx path/to/file.xlsx --info

# 转换为 CSV 或 JSON 格式
bun run src/index.ts xlsx-convert path/to/file.xlsx -o output.csv
bun run src/index.ts xlsx-convert path/to/file.xlsx -o output.json

# 读取并显示 DOCX 内容
bun run src/index.ts docx path/to/file.docx

# 显示段落数
bun run src/index.ts docx path/to/file.docx --info
```

## 项目结构 | Project Structure

```
bun-cli/
├── src/
│   ├── index.ts              # 入口文件（CLI/UI 双模式分发）
│   ├── config/
│   │   ├── app.ts            # 应用全局配置
│   │   └── user-settings.ts  # 用户设置持久化管理
│   ├── i18n/
│   │   └── index.ts          # 国际化翻译表（4 种语言）
│   └── module/
│       ├── cli.ts            # CLI 命令行实现（commander）
│       ├── docx-engine.ts    # DOCX 读写引擎
│       ├── xlsx-engine.ts    # XLSX 读写引擎（SheetJS 封装）
│       ├── worker.ts         # HTTP 服务端 + 内嵌前端 SPA
│       └── generated-assets.ts # 自动生成的资源映射表
├── scripts/
│   ├── embed-assets.ts       # 资源嵌入工具（图片→Base64 TS）
│   └── clean.ts              # 清理脚本
├── tests/                    # 测试文件与截图参考
├── assets/                   # 静态资源（图标等）
├── build.ts                  # 跨平台编译脚本
├── package.json
└── tsconfig.json             # TypeScript 配置（含路径别名）
```

## 编译构建 | Build

```bash
# 完整编译流程（嵌入资源 → 编译二进制）
bun run compile

# 仅编译（需先执行 bun run compile-assets）
bun run build

# 清理构建产物
bun run clean

# 嵌入资源文件
bun run compile-assets
```

编译产物输出至 `bin/` 目录。Windows 平台会自动注入图标和版本信息。

Build outputs to the `bin/` directory. On Windows, icons and version info are automatically injected via `rcedit`.

## 测试 | Test

```bash
bun test
```

测试覆盖范围：

| 测试文件 | 覆盖模块 |
|---------|---------|
| `xlsx-engine.test.ts` | XLSX 读写、样式提取、CRUD、合并单元格、规则应用 |
| `docx-engine.test.ts` | DOCX 读写、段落操作、样式保留、往返测试 |
| `cli.test.ts` | CLI 命令行参数解析 |
| `i18n.test.ts` | 国际化翻译完整性 |
| `user-settings.test.ts` | 用户设置持久化 |

## 技术栈 | Tech Stack

| 类别 | 技术 |
|-----|------|
| 运行时 | [Bun](https://bun.sh/) |
| XLSX 引擎 | [SheetJS (xlsx)](https://sheetjs.com/) |
| DOCX 引擎 | [docx](https://docx.js.org/) |
| GUI | [webview-bun](https://github.com/nicedayzhu/webview-bun) |
| CLI | [commander](https://commander.js/) |
| 语言 | TypeScript |

## API 接口 | API Endpoints

GUI 模式下，内置 HTTP 服务器提供以下 REST API：

### XLSX 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/xlsx/open` | 打开 XLSX 文件 |
| POST | `/api/xlsx/save` | 保存/导出 XLSX |
| POST | `/api/xlsx/cell` | 更新单元格值 |
| POST | `/api/xlsx/apply-rules` | 应用用户自定义规则 |
| POST | `/api/xlsx/merge` | 合并单元格 |
| POST | `/api/xlsx/unmerge` | 取消合并单元格 |
| POST | `/api/xlsx/add-row` | 添加行 |
| POST | `/api/xlsx/add-col` | 添加列 |
| POST | `/api/xlsx/delete-row` | 删除行 |
| POST | `/api/xlsx/delete-col` | 删除列 |
| POST | `/api/xlsx/add-sheet` | 新建工作表 |
| POST | `/api/xlsx/delete-sheet` | 删除工作表 |
| POST | `/api/xlsx/rename-sheet` | 重命名工作表 |

### DOCX 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/docx/open` | 打开 DOCX 文件 |
| POST | `/api/docx/save` | 保存/导出 DOCX |
| POST | `/api/docx/add-paragraph` | 添加段落 |
| POST | `/api/docx/update-paragraph` | 更新段落 |
| POST | `/api/docx/delete-paragraph` | 删除段落 |

### 设置与国际化

| 方法 | 路径 | 说明 |
|------|------|------|
| GET/POST | `/api/settings` | 读取/保存用户设置 |
| GET | `/api/i18n` | 获取国际化消息 |

## License

[MIT](LICENSE)

## Author

[VDHewei](https://github.com/VDHewei)
