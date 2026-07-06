根据 Excel 映射表批量重命名指定目录下的文件，支持递归查找子目录、文件备份、加密 Excel 文件读取等功能。

## 功能特性

- 根据 Excel 映射表批量重命名文件
- 支持递归查找子目录中的文件
- 支持文件备份功能
- 支持加密 Excel 文件（通过 WPS/Office COM 组件自动解密）
- 所有路径使用相对路径，便于移动和分享
- 详细的调试日志

## 版本信息

- **版本**: 1.1.5
- **作者**: CHEN
- **联系方式**: 13516111231
- **创建日期**: 2026-03-13
- **最后更新**: 2026-03-23

## 系统要求

- Python 3.6+
- Windows 系统（COM 组件功能需要）
- WPS 或 Microsoft Office（用于读取加密 Excel 文件）

## 安装

### 1. 克隆项目

```bash
git clone https://github.com/yourusername/Batch_Renamer.git
cd Batch_Renamer
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

或使用启动脚本自动安装：

```bash
# Windows
启动.bat
```

## 使用方法

### 方法一：使用启动脚本（推荐）

双击 `启动.bat` 运行程序。

### 方法二：命令行运行

```bash
python batch_rename.py
```

### 命令行参数

```bash
python batch_rename.py [参数]
```

**常用参数：**

- `--config, -c` - 指定配置文件路径（默认：config.ini）
- `--mapping, -m` - 指定映射表文件路径
- `--root-dir, -r` - 指定根目录
- `--no-header` - 映射表没有标题行
- `--execute` - 直接执行重命名而不询问
- `--debug` - 启用调试模式
- `--recursive-find` - 递归查找文件（在子目录中查找）
- `--version, -v` - 显示版本信息

**示例：**

```bash
# 使用默认配置
python batch_rename.py

# 指定映射表
python batch_rename.py --mapping mylist.xlsx

# 启用调试模式
python batch_rename.py --debug

# 递归查找文件
python batch_rename.py --recursive-find
```

## 配置文件说明

编辑 `config.ini` 自定义设置：

```ini
[Paths]
; 映射表文件路径（支持相对路径）
mapping_file = rename.xlsx
; 根目录路径（用于解析相对路径，支持相对路径）
root_dir = .
; 备份目录路径（支持相对路径）
backup_dir = backup

[Settings]
; 映射表是否有标题行
has_header = true
; 是否自动执行（不询问）
auto_execute = false
; 是否清理文件名中的非法字符
sanitize_filename = true
; 非法字符替换符
illegal_char_replacement = _
; 是否创建文件备份
create_backup = false
; 路径列索引（从 0 开始）
path_column = 0
; 原文件名列索引（从 0 开始）
oldname_column = 1
; 新文件名列索引（从 0 开始）
newname_column = 2
; 是否启用调试模式
debug_mode = true
; 是否使用 COM 组件读取加密文件
use_com_for_encrypted = true
; COM 应用程序选择：auto, wps, excel
com_application = auto

[Logging]
; 是否启用日志记录
enable_logging = true
; 日志文件路径（支持相对路径）
log_file = rename_log.txt
; 日志级别：DEBUG, INFO, WARNING, ERROR
log_level = DEBUG
```

## 映射表格式

Excel 文件（`rename.xlsx`）必须包含以下三列：

| 文件所在路径 | 原始名称 | 改名 | 备注 |
|--------------|----------|------|------|
| D:\文件夹 1   | 旧文件 1.pdf | 新文件 1.pdf | 可选 |
| D:\文件夹 2   | 旧文件 2.pdf | 新文件 2.pdf | 可选 |

**注意事项：**

1. 第一行可以是标题行（根据 `has_header` 设置）
2. 路径可以是绝对路径，也可以是相对于 `root_dir` 的相对路径
3. 如果文件不在指定路径下，可以使用 `--recursive-find` 参数递归查找
4. 列索引可在配置文件中调整

## 操作流程

1. **准备映射表**：在 Excel 中填写文件路径、原始名称和新名称
2. **检查配置**：根据需要修改 `config.ini` 中的设置
3. **运行程序**：双击 `启动.bat` 或在命令行中运行
4. **预览结果**：程序会显示将要重命名的文件列表
5. **确认执行**：输入 `y` 确认重命名，输入 `n` 取消
6. **查看日志**：重命名完成后可查看 `rename_log.txt` 了解详细过程

## 注意事项

1. **备份重要文件**：建议先在小范围测试，或开启 `create_backup` 选项
2. **文件名冲突**：如果新文件名已存在，该文件会被跳过
3. **非法字符**：Windows 文件名不能包含 `\/:*?"<>|`，程序会自动替换
4. **路径格式**：建议使用绝对路径，或确保相对路径正确
5. **Excel 格式**：只支持 `.xlsx` 格式，需要安装 `openpyxl`

## 常见问题

### Q: 运行后提示 "找不到映射表文件"？

**A**: 检查 `rename.xlsx` 是否在当前目录，或使用 `--mapping` 参数指定路径

### Q: 提示 "文件不存在"？

**A**: 检查 Excel 中的路径和文件名是否正确，或使用 `--recursive-find` 参数

### Q: 如何查看详细错误信息？

**A**: 在 `config.ini` 中设置 `debug_mode = true`，查看 `rename_log.txt`

### Q: 可以批量重命名子文件夹中的文件吗？

**A**: 可以，使用 `--recursive-find` 参数，程序会在子目录中递归查找

### Q: 如何处理加密的 Excel 文件？

**A**: 程序会自动检测加密文件，并通过 WPS/Office COM 组件自动解密。需要：
- 安装 `pywin32`：`pip install pywin32`
- 系统已安装 WPS 或 Microsoft Office
- 配置文件中 `use_com_for_encrypted = true`

## 项目结构

```
Batch_Renamer/
├── batch_rename.py      # 主程序
├── config.ini           # 配置文件
├── requirements.txt     # Python 依赖
├── .gitignore          # Git 忽略文件
├── README.md           # 项目说明
└── 启动.bat             # Windows 启动脚本
```

## 许可证

本项目仅供学习和内部使用。

## 联系方式

如有问题或建议，请联系：

- **作者**：CHEN
- **电话**：13516111231

---

**感谢使用！**
