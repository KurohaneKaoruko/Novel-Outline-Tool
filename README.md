# novel-outline-tool

Windows 原生轻量桌面小说大纲工具（Rust / Win32），以“项目文件夹 + Markdown”方式组织与编辑小说资料。

## 特性

- 原生 Win32 GUI（不依赖 WebView / Electron）
- 章节、角色、世界观、时间线模块化管理（均为 Markdown 文件）
- 章节拖拽排序（会自动重编号文件名序号）
- 搜索过滤、深浅色主题
- 自动保存与 `.backup/` 快照备份
- 导出项目为独立文件夹
- 可选在线更新检查（在 `project.md` 配置 `update_url`）

## 环境要求

- Windows（强依赖 Win32/WinHTTP/RichEdit）
- Rust 1.75+（见 [Cargo.toml](file:///d:/.Programs/.Program.Project/novel-outline-tool/Cargo.toml)）

## 从源码构建

```bash
cargo build --release
```

产物位于：

- `target/release/novel-outline-tool.exe`

开发运行：

```bash
cargo run
```

## 使用方式（项目文件夹）

该工具操作的对象是一个“项目文件夹”。首次打开时会初始化如下结构：

```text
your-project/
  project.md
  chapters/
  characters/
  world/
  timeline/
  exports/
  .backup/
```

### project.md（YAML front matter）

`project.md` 采用 YAML front matter 保存元信息（示例）：

```markdown
---
name: 我的小说
created_unix: 1700000000
format_version: 1
theme: dark
left_pane_ratio: 0.33
update_url: https://example.com/novel-outline-tool/update.json
---

这里是项目说明正文（可选）。
```

常用字段：

- `name`：项目名称
- `format_version`：格式版本
- `theme`：主题（如 `light`/`dark`，以程序实际支持为准）
- `left_pane_ratio`：左侧面板比例（0~1）
- `update_url`：更新信息 JSON 地址（仅支持 HTTPS）

### chapters/（章节文件）

- 章节以 Markdown 文件存放于 `chapters/`
- 文件名带序号前缀（例如 `0001-第一章.md`）
- 拖拽排序后会重命名并自动重编号

### .backup/（自动备份）

工具会在自动保存时写入快照到 `.backup/`，并定期清理旧备份（默认保留最近 30 份）。

### exports/（导出）

导出会生成一个独立的 `*-export-*` 目录（位于你选择的导出位置），其中包含当前项目文件的副本，便于分享或归档。

## 代码结构（开发者）

- 入口与 UI：`src/main.rs`
- 数据模型：`src/domain/`
- 项目落盘/备份/原子写：`src/storage/`
- 在线更新检查：`src/update.rs`
- 资源嵌入：`resources/` + `build.rs`

## 已知限制

- 仅支持 Windows
- 更新检查仅支持 HTTPS

## License

本项目采用 GPL-3.0 许可证，详见 [LICENSE](file:///d:/.Programs/.Program.Project/novel-outline-tool/LICENSE)。
