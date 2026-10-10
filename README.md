# 字段同步工具

用于关联性强、且具有多映射关系的多表格管理与配置。

一个基于 Python + openpyxl + tkinter 的 xlsx 字段映射管理工具：用一张「工具表」登记多个表格之间的同义字段（映射组），然后在图形界面里完成跨表字段同步和字段值翻译。

## 功能

- **同步**：以一个表格的字段值为标准，按键字段（映射组）对齐行，把值同步到其他表格的同义字段
- **字段替换**：把某列中填的 A 类值（如名称）按参照表批量翻译成 B 类值（如 ID）
- **登记字段**：文件夹树 → 子表 → 字段逐级登记，支持任意层级子目录，支持新建映射组
- **防护**：启动校验（工具表格式 + 映射有效性）、操作前预览、写入前自动备份、扫描缓存加速
- 表头永不被修改，所有操作只写数据区

## 使用

```bash
pip install openpyxl
python sync_gui.py          # 或双击 字段同步工具.vbs（Windows 无黑窗启动）
```

也可以直接在 [Releases](../../releases) 页面下载打包好的 exe（免装 Python）和工具表模板。

首次启动会引导你选择或新建一个工具表（含「总表」「字段表」「使用说明」三个 sheet）。工具表和被管理的 xlsx 放在同一目录（支持子目录）。

## 打包为 exe（可选）

```bash
pip install pyinstaller
python -m PyInstaller --onefile --noconsole --clean --name 字段同步工具 \
  --exclude-module PIL --exclude-module lxml --exclude-module ssl \
  --exclude-module unittest --exclude-module doctest --exclude-module pydoc \
  --exclude-module turtle --exclude-module lib2to3 --exclude-module xmlrpc \
  --exclude-module multiprocessing --exclude-module concurrent sync_gui.py
```

## 文件说明

| 文件 | 说明 |
|---|---|
| `sync_gui.py` | 工具主程序（纯逻辑函数 + tkinter 界面） |
| `字段同步工具.vbs` | Windows 无控制台启动器 |
