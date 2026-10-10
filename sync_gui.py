# -*- coding: utf-8 -*-
"""字段映射同步工具（窗口版）

工具表.xlsx 结构:
  总表:   文件名称 | 表格名称            —— 登记表格清单
  字段表: 映射组 | 文件名称 | 表格名称 | 字段名  —— 字段映射缓存(长表)
          同一映射组里的字段互为映射，按需登记，没登记的不参与同步。

运行: python sync_gui.py
窗口中: 选 映射组(键) → 数据源表格 → 同步映射组(多选) → 被同步表格(多选)，
先「预览」再「开始同步」。「登记字段」可把文件夹里的 xlsx 表格及其字段
加入指定映射组（可新建映射组）。

同步规则:
  · 按键组的值匹配行，把数据源的值写入被同步表格在同一映射组里的字段
  · 目标表格中所有键值相同的行都会更新
  · 数据源中该值为空的行不同步（不会清空目标数据）
  · 各表格表头都应在第 1 行
"""
import json
import os
import shutil
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk, scrolledtext, messagebox, filedialog

import openpyxl
from openpyxl.styles import Font, PatternFill

TOOL_FILE = "字段同步工具表.xlsx"
MAIN_SHEET = "总表"
FIELD_SHEET = "字段表"
CONFIG_FILE = "工具配置.json"
BACKUP_DIR = "字段同步备份"
MAIN_HEADERS = ["文件名称", "表格名称"]
FIELD_HEADERS = ["映射组", "文件名称", "表格名称", "字段名"]
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")

INSTRUCTIONS = [
    ("字段映射同步工具 - 使用说明", True),
    ("", False),
    ("一、表格结构", True),
    ("「总表」: 登记表格清单, 只有两列 —— 文件名称 | 表格名称", False),
    ("「字段表」: 字段映射缓存(长表), 四列 —— 映射组 | 文件名称 | 表格名称 | 字段名", False),
    ("· 同一个映射组里的字段互为映射(如「作物名称」组里有 作物名称/茶树名称/茶树名)", False),
    ("· 按需缓存: 只登记你关心的字段, 没登记的字段不参与同步", False),
    ("· 数据文件按【工具表所在目录】查找, 支持任意层级的子目录, 文件名称显示相对路径(如 子目录/配置表.xlsx)", False),
    ("· 备份目录(默认「字段同步备份」)下的文件不会被扫描", False),
    ("· 含有「总表」「字段表」的文件会被识别为工具表, 不会当作数据表扫描", False),
    ("", False),
    ("二、同步操作", True),
    ("1. 打开工具(首次启动会引导你选择或新建一个表格作为工具表)", False),
    ("2. 选择「同步映射组」: 要同步的字段组(如 作物名称)", False),
    ("3. 选择「同步源字段」: 组内各字段已标明位置(文件|表格|字段名), 选一个作为数值标准", False),
    ("4. 在「被同步字段」里勾选目标(可点「全选」选中组内其余所有字段)", False),
    ("5. 选择「键映射组」: 用来对齐各表数据行的字段组(如 作物ID), 候选已按所选字段自动筛选", False),
    ("6. 点「预览」确认改动, 再点「开始同步」(默认自动备份)", False),
    ("· 备份目录默认为工具表旁的「字段同步备份」文件夹, 可点「设置备份目录」按钮更改, 选择会被记住", False),
    ("", False),
    ("三、登记字段(在工具窗口内完成)", True),
    ("1. 点「登记字段」按钮, 按 文件 → 子表 → 未登记字段 逐级选择(字段列表只显示还没登记的)", False),
    ("2. 勾选字段(可输入关键字筛选), 选择或输入一个映射组(可新建), 点「加入」", False),
    ("3. 一个表格可以分多次把不同字段加入不同映射组, 最后点「保存」写入工具表", False),
    ("· 也可以直接在「总表」「字段表」里手动加行, 效果相同; 注意同一表格在同一映射组里只能有一个字段, 重复添加视为修改(替换)", False),
    ("· 删除整个映射组: 点主窗口「删除映射组」按钮, 选组后确认(只删映射关系, 不动数据表)", False),
    ("", False),
    ("四、同步规则", True),
    ("· 按键组的值匹配行, 把数据源的值写入被同步表格在同一映射组里的字段", False),
    ("· 目标表格中所有键值相同的行都会更新", False),
    ("· 数据源中该值为空的行不同步(不会清空目标数据)", False),
    ("· 各表格的表头都应在第 1 行", False),
    ("· 所有操作都只改写数据区(第 2 行起), 不会修改任何表格的表头", False),
    ("· 公式保护: 源值是公式的不同步/不参照, 目标单元格是公式的不覆盖, 均会在日志给出警告", False),
    ("", False),
    ("五、字段替换(页签2)", True),
    ("共用: 「被替换字段」逐级选择 文件→子表→字段(所有字段都能选, 含未登记的; 可多选, 跨子表累积);", False),
    ("     输入关键字时切换为完整路径平铺筛选", False),
    ("方式一「按参照表(映射组)」: 声明 当前值是/替换成 两个映射组, 选一张同时登记了两组的表当字典;", False),
    ("  被替换列里的当前值去参照表查, 找到就替换。适合『ID列里临时填了名称, 批量翻回ID』", False),
    ("方式二「按字段匹配」: 选参照表格 + 匹配字段 + 取值字段; 按目标表同行的匹配字段值去参照表对齐行,", False),
    ("  把取值字段的值写入被替换列。目标子表需有与匹配字段同名的列", False),
    ("· 查不到对照的保留原样并列入日志; 空单元格跳过; 重复对照取第一行并警告", False),
    ("", False),
    ("六、常见问题", True),
    ("", False),
    ("Q1: 键ID字段含义相同、但在不同表里命名不同(如 作物ID / 茶树ID), 能用来匹配吗?", True),
    ("可以, 而且这就是映射组的设计用途。同步时每个表格用自己在键映射组里登记的字段名取键值,", False),
    ("再按键值匹配行, 字段名不同不影响。例如种植表的 作物ID=202001 和茶树表的 茶树ID=202001", False),
    ("会被认成同一行。", False),
    ("注意: 遇到还没归组的同义字段, 用「登记字段」把它加进【已有的】映射组, 不要新建组,", False),
    ("否则新旧字段就断开映射了。", False),
    ("", False),
    ("Q2: 键映射组误选了非ID类组(如 培育时间)会有什么影响?", True),
    ("不会报错, 但结果很可能是错的。键值不唯一时: 数据源里同键值只取第一行, 目标表里所有同键值", False),
    ("的行会被全部刷成那一行的值。判断标准: 该字段在每个参与表格里是否每行唯一(如 文本ID 就可以当键)。", False),
    ("防护: 同步前先点「预览」, 改动行数远超预期通常就是键选错了; 默认写入前会自动备份。", False),
]


def app_dir():
    """脚本/exe 自身所在目录（工具表和被管理的 xlsx 都放在这里）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def load_config(base_dir):
    """读取工具配置（工具表路径、备份目录等），没有或损坏时返回空 dict。"""
    cfg = base_dir / CONFIG_FILE
    if not cfg.exists():
        return {}
    try:
        return json.loads(cfg.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(base_dir, **kw):
    """更新工具配置（只覆盖传入的键）。"""
    data = load_config(base_dir)
    data.update(kw)
    (base_dir / CONFIG_FILE).write_text(
        json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _config_path(base_dir, value):
    """配置里的路径: 相对路径相对 app 目录解析。"""
    p = Path(value)
    if not p.is_absolute():
        p = base_dir / p
    return p


def find_tool_config(base_dir):
    """读取上次选择的工具表路径，没有配置或文件已不存在时返回 None。"""
    v = load_config(base_dir).get("tool_file", "")
    if not v:
        return None
    p = _config_path(base_dir, v)
    return p if p.exists() else None


def save_tool_config(base_dir, path):
    try:
        s = str(path.relative_to(base_dir))
    except ValueError:
        s = str(path)
    save_config(base_dir, tool_file=s)


def find_backup_config(base_dir):
    """读取用户设置的备份目录，未设置时返回 None。"""
    v = load_config(base_dir).get("backup_dir", "")
    return _config_path(base_dir, v) if v else None


def save_backup_config(base_dir, path):
    try:
        s = str(path.relative_to(base_dir))
    except ValueError:
        s = str(path)
    save_config(base_dir, backup_dir=s)


def check_tool_format(path):
    """检查表格: 'ok'=符合工具表格式, 'empty'=空表(可初始化), 'invalid'=两者都不是。"""
    try:
        wb = openpyxl.load_workbook(path)
    except Exception:
        return "invalid"
    try:
        has_data = any(
            any(c.value is not None for c in row)
            for ws in wb.worksheets
            for row in ws.iter_rows(max_row=1)
        ) or any(ws.max_row > 1 for ws in wb.worksheets)
        if not has_data:
            return "empty"
        if MAIN_SHEET in wb.sheetnames and FIELD_SHEET in wb.sheetnames:
            h1 = [norm(c.value) for c in wb[MAIN_SHEET][1][:2]]
            h2 = [norm(c.value) for c in wb[FIELD_SHEET][1][:4]]
            if h1 == MAIN_HEADERS and h2 == FIELD_HEADERS:
                return "ok"
        return "invalid"
    finally:
        wb.close()


def init_tool_file(path):
    """把表格初始化为工具表: 写入使用说明, 确保「总表」「字段表」及其表头存在。"""
    path = Path(path)
    if path.exists():
        wb = openpyxl.load_workbook(path)
    else:
        wb = openpyxl.Workbook()
        wb.active.title = "使用说明"
    if "使用说明" in wb.sheetnames:
        del wb["使用说明"]
    ws = wb.create_sheet("使用说明", 0)
    for i, (text, bold) in enumerate(INSTRUCTIONS, 1):
        c = ws.cell(i, 1, text)
        if bold:
            c.font = Font(bold=True, size=12 if "、" not in text else 11)
    ws.column_dimensions["A"].width = 95
    for name, headers, widths in [
        (MAIN_SHEET, MAIN_HEADERS, [26, 24]),
        (FIELD_SHEET, FIELD_HEADERS, [16, 26, 24, 18]),
    ]:
        if name in wb.sheetnames:
            ws = wb[name]
        else:
            ws = wb.create_sheet(name)
        for j, h in enumerate(headers, 1):
            c = ws.cell(1, j, h)
            c.font = Font(bold=True)
            c.fill = HEADER_FILL
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    wb.save(path)


def norm(v):
    """把单元格值归一化成可比较的字符串（1 / 1.0 / '1' 视为相同）。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def load_tool(path):
    """读取工具表，返回 (tables, groups)。
    tables: [(文件, 表格), ...]（保持总表顺序）
    groups: {映射组: {(文件, 表格): 字段名}}"""
    wb = openpyxl.load_workbook(path, read_only=True)
    if MAIN_SHEET not in wb.sheetnames or FIELD_SHEET not in wb.sheetnames:
        wb.close()
        raise ValueError(f"工具表需要包含「{MAIN_SHEET}」和「{FIELD_SHEET}」两个工作表")
    tables = []
    for r in wb[MAIN_SHEET].iter_rows(min_row=2, values_only=True):
        if r and norm(r[0] if len(r) > 0 else None) and norm(r[1] if len(r) > 1 else None):
            tables.append((norm(r[0]), norm(r[1])))
    groups = {}
    for r in wb[FIELD_SHEET].iter_rows(min_row=2, values_only=True):
        if not r or len(r) < 4:
            continue
        g, f, s, col = (norm(r[0]), norm(r[1]), norm(r[2]), norm(r[3]))
        if g and f and s and col:
            groups.setdefault(g, {})[(f, s)] = col
    wb.close()
    return tables, groups


def valid_headers(ws):
    hdr = list(next(ws.iter_rows(max_row=1, values_only=True), None) or [])
    while hdr and hdr[-1] is None:
        hdr.pop()
    if not hdr or not all(isinstance(h, str) and h.strip() for h in hdr):
        return None
    if len(set(hdr)) != len(hdr):
        return None
    return [h.strip() for h in hdr]


SCAN_CACHE_FILE = "扫描缓存.json"


def scan_tables(base_dir, tool_path=None, exclude_dir=None, cache_dir=None):
    """扫描根目录（含任意层级子目录）下所有可用的 xlsx 表格。
    返回 [(相对路径, [(子表名, [表头...]), ...]), ...]，按文件分组。
    exclude_dir（备份目录）下的文件不参与。
    cache_dir 提供时启用扫描缓存（按 路径+修改时间+大小 判断变化），
    未变化的文件不重复解析，大幅加快多文件场景的扫描。"""
    result = []
    base_dir = Path(base_dir)
    tool_resolved = Path(tool_path).resolve() if tool_path else None
    excl_resolved = Path(exclude_dir).resolve() if exclude_dir else None
    cache = {}
    cache_path = Path(cache_dir) / SCAN_CACHE_FILE if cache_dir else None
    if cache_path and cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            cache = {}
    new_cache = {}
    for p in sorted(base_dir.rglob("*.xlsx")):
        if p.name.startswith("~$"):
            continue
        rp = p.resolve()
        if tool_resolved and rp == tool_resolved:
            continue
        if tool_resolved is None and p.name == TOOL_FILE:
            continue
        if excl_resolved and excl_resolved in rp.parents:
            continue
        try:
            st = p.stat()
        except OSError:
            continue
        key = str(rp)
        ent = cache.get(key)
        if ent and ent.get("mtime") == st.st_mtime and ent.get("size") == st.st_size:
            new_cache[key] = ent
        else:
            ent = {"mtime": st.st_mtime, "size": st.st_size, "tool": False, "sheets": []}
            try:
                wb = openpyxl.load_workbook(p, read_only=True)
            except Exception:
                new_cache[key] = ent
                continue
            if MAIN_SHEET in wb.sheetnames and FIELD_SHEET in wb.sheetnames:
                ent["tool"] = True  # 工具表格式的文件不作为数据表
            else:
                for ws in wb.worksheets:
                    hdr = valid_headers(ws)
                    if hdr:
                        ent["sheets"].append([ws.title, hdr])
            wb.close()
            new_cache[key] = ent
        if ent["tool"] or not ent["sheets"]:
            continue
        result.append((p.relative_to(base_dir).as_posix(), ent["sheets"]))
    if cache_path:
        try:
            cache_path.write_text(json.dumps(new_cache, ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass
    return result


def build_file_tree(tree, rel_paths):
    """把相对路径列表建成 文件夹/文件 树。文件夹节点 iid='dir:路径', 文件叶子 iid='file:路径'。"""
    nodes = {}
    for rel in sorted(rel_paths):
        parts = rel.split("/")
        parent = ""
        for i, part in enumerate(parts[:-1]):
            key = "/".join(parts[:i + 1])
            if key not in nodes:
                nodes[key] = tree.insert(parent, "end", iid=f"dir:{key}", text=part)
            parent = nodes[key]
        tree.insert(parent, "end", iid=f"file:{rel}", text=parts[-1])


def tree_selected_file(tree):
    """返回树中选中的文件相对路径; 选中文件夹或未选中返回 None。"""
    sel = tree.selection()
    if sel and sel[0].startswith("file:"):
        return sel[0][5:]
    return None


def tree_selected_prefix(tree):
    """返回树中选中的路径前缀: 文件叶子返回该文件路径, 文件夹返回 '文件夹/' 前缀。"""
    sel = tree.selection()
    if not sel:
        return None
    iid = sel[0]
    if iid.startswith("file:"):
        return iid[5:]
    return iid[4:] + "/"


def add_associations(tool_path, additions):
    """把新登记写入工具表。additions: [(文件, 表格, 字段名, 映射组), ...]。
    同一映射组内同一表格只保留一个字段：重复添加视为修改（替换字段名）。
    返回 (新增行数, 更新行数)。"""
    wb = openpyxl.load_workbook(tool_path)
    ws_main, ws_field = wb[MAIN_SHEET], wb[FIELD_SHEET]
    registered = set()
    for r in ws_main.iter_rows(min_row=2, values_only=True):
        if r and norm(r[0] if len(r) > 0 else None) and norm(r[1] if len(r) > 1 else None):
            registered.add((norm(r[0]), norm(r[1])))
    existing = {}  # (映射组, 文件, 表格) -> 行号
    for i, r in enumerate(ws_field.iter_rows(min_row=2, values_only=True), start=2):
        if r and len(r) >= 4 and norm(r[0]) and norm(r[1]) and norm(r[2]):
            existing[(norm(r[0]), norm(r[1]), norm(r[2]))] = i
    added = updated = 0
    for f, s, col, g in additions:
        if (f, s) not in registered:
            ws_main.append([f, s])
            registered.add((f, s))
        key = (g, f, s)
        if key in existing:
            row_i = existing[key]
            if norm(ws_field.cell(row_i, 4).value) != col:
                ws_field.cell(row_i, 4).value = col
                updated += 1
        else:
            ws_field.append([g, f, s, col])
            existing[key] = ws_field.max_row
            added += 1
    wb.save(tool_path)
    return added, updated


def check_mappings(base_dir, tables, groups, cache_dir=None, tool_path=None, exclude_dir=None):
    """检查字段表里的映射是否仍然有效（文件/工作表/字段列都存在）。
    返回失效映射列表 [(映射组, 文件, 表格, 字段名, 原因), ...]。
    表头信息走 scan_tables 的扫描缓存，多文件场景下更快。"""
    lost = []
    registered = set(tables)
    scanned = {
        f: {s: hdr for s, hdr in sheets}
        for f, sheets in scan_tables(base_dir, tool_path, exclude_dir, cache_dir)
    }
    for g, mapping in groups.items():
        for (f, s), col in mapping.items():
            if (f, s) not in registered:
                lost.append((g, f, s, col, "表格未在总表登记"))
                continue
            if not (base_dir / f).exists():
                lost.append((g, f, s, col, "文件不存在"))
                continue
            info = scanned.get(f)
            if not info:
                lost.append((g, f, s, col, "文件无法读取"))
            elif s not in info:
                lost.append((g, f, s, col, "工作表不存在"))
            elif col not in info[s]:
                lost.append((g, f, s, col, "字段列不存在"))
    return lost


def remove_mappings(tool_path, lost):
    """从字段表删除失效映射行，返回删除条数。"""
    targets = {(g, f, s, col) for g, f, s, col, _ in lost}
    wb = openpyxl.load_workbook(tool_path)
    ws = wb[FIELD_SHEET]
    n = 0
    for r in range(ws.max_row, 1, -1):
        vals = tuple(norm(ws.cell(r, c).value) for c in range(1, 5))
        if vals in targets:
            ws.delete_rows(r)
            n += 1
    wb.save(tool_path)
    return n


def delete_group(tool_path, group):
    """删除整个映射组（字段表中该组的所有行），返回删除条数。"""
    wb = openpyxl.load_workbook(tool_path)
    ws = wb[FIELD_SHEET]
    n = 0
    for r in range(ws.max_row, 1, -1):
        if norm(ws.cell(r, 1).value) == group:
            ws.delete_rows(r)
            n += 1
    wb.save(tool_path)
    return n


def save_with_backup(wb, rel, base_dir, backup_dir, stamp):
    """保存工作簿: 先把磁盘上的原文件备份一次, 再写入临时文件并原子替换原文件。
    即使保存中途失败, 原文件也不会损坏。返回备份文件名（未备份返回 None）。"""
    tp = base_dir / rel
    bak_name = None
    if backup_dir:
        bdir = Path(backup_dir)
        bdir.mkdir(parents=True, exist_ok=True)
        safe = "__".join(Path(rel).with_suffix("").parts)
        bak_name = f"{safe}.{stamp}{tp.suffix}"
        shutil.copy2(tp, bdir / bak_name)  # 此刻磁盘上仍是原始文件
    tmp = tp.with_name(tp.name + ".tmp")
    wb.save(tmp)
    os.replace(tmp, tp)
    return bak_name


def is_formula_like(v):
    """判断单元格值是否是公式（或以 = 开头的文本，写入后会变成公式，有注入风险）。"""
    return isinstance(v, str) and v.startswith("=")


def header_map_of(ws):
    """读取已打开工作表的第 1 行表头，返回 {列名: 列号(1-based)}。"""
    hmap = {}
    for c in ws[1]:
        h = norm(c.value)
        if h and h not in hmap:
            hmap[h] = c.column
    return hmap


def read_sheet(path, sheet):
    """打开目标表格，返回 (workbook, worksheet, {列名: 列号})。"""
    wb = openpyxl.load_workbook(path)
    if sheet not in wb.sheetnames:
        raise KeyError(f"工作表「{sheet}」不存在于 {Path(path).name}")
    ws = wb[sheet]
    return wb, ws, header_map_of(ws)


def run_sync(base_dir, groups, key_group, src, sync_groups, targets, backup_dir, dry, log):
    """执行同步（dry=True 只预览不写盘）。src/targets 为 (文件, 表格)。
    backup_dir 为 None 时不备份。返回改动单元格数。
    日志按统一模板输出: 标题+配置 → 明细 → 警告 → 执行结果 → 汇总。"""
    _d, _w = [], []
    _log_out = log

    def log(msg):
        (_w if msg.strip().startswith("[警告]") else _d).append(msg)

    def flush():
        if _d or _w:
            _log_out("\n".join(_d + _w))
            _d.clear()
            _w.clear()

    tot_blocked = 0    # 公式拦截数
    tot_unmatched = 0  # 键值未匹配数

    src_key_col = groups.get(key_group, {}).get(src)
    if not src_key_col:
        log(f"[错误] 数据源表格在映射组「{key_group}」里没有登记字段")
        flush()
        return 0
    try:
        swb, sws, shdr = read_sheet(base_dir / src[0], src[1])
    except Exception as ex:
        log(f"[错误] 读取数据源失败: {ex}")
        flush()
        return 0
    if src_key_col not in shdr:
        log(f"[错误] 数据源表头里找不到键列「{src_key_col}」")
        return 0

    src_fields = {}
    for g in sync_groups:
        col = groups.get(g, {}).get(src)
        if not col:
            log(f"[跳过] 数据源表格在映射组「{g}」里没有登记字段")
            continue
        if col not in shdr:
            log(f"[跳过] 数据源表头里找不到字段列「{col}」(映射组「{g}」)")
            continue
        src_fields[g] = shdr[col]
    if not src_fields:
        log("[中止] 数据源没有任何可同步的字段")
        flush()
        return 0

    key_c = shdr[src_key_col]
    src_data = {}
    dup_keys = 0
    formula_src = []  # (行号, 映射组) 源值是公式的
    for r in range(2, sws.max_row + 1):
        k = norm(sws.cell(r, key_c).value)
        if not k:
            continue
        if k in src_data:
            dup_keys += 1
            continue
        vals = {}
        for g, c in src_fields.items():
            v = sws.cell(r, c).value
            if is_formula_like(v):
                formula_src.append((r, g))
                v = None  # 公式不同步，视同空值跳过
            vals[g] = v
        src_data[k] = vals
    swb.close()
    if formula_src:
        sample = ", ".join(f"行{r}「{g}」" for r, g in formula_src[:5])
        log(f"[警告] 数据源有 {len(formula_src)} 个值是公式，已跳过（工具不同步公式内容）: {sample}"
            + (" ……" if len(formula_src) > 5 else ""))
    if dup_keys:
        log(f"[提示] 数据源有 {dup_keys} 个重复的键值，已取首次出现的行")
    log(f"数据源: {src[0]} | {src[1]}，共 {len(src_data)} 个键值")
    log(f"键映射组: {key_group} | 同步映射组: {', '.join(src_fields)}")

    total_changes = 0
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    # 按文件分批: 每个文件 打开一次 → 校验并生成变更计划 → 备份一次 → 应用全部修改 → 原子保存
    by_file = {}
    for tgt in targets:
        if tgt == src:
            continue
        by_file.setdefault(tgt[0], []).append(tgt)

    results = []
    for f, tgts in by_file.items():
        try:
            twb = openpyxl.load_workbook(base_dir / f)
        except Exception as ex:
            log(f"[错误] {f}: 打开失败: {ex}")
            results.append(f"{f}: 打开失败（{ex}）")
            continue
        file_changes = 0
        for tgt in tgts:
            s = tgt[1]
            label = f"{f} | {s}"
            tgt_key_col = groups.get(key_group, {}).get(tgt)
            if not tgt_key_col:
                log(f"[跳过] {label}: 在映射组「{key_group}」里没有登记字段")
                continue
            if s not in twb.sheetnames:
                log(f"[错误] {label}: 工作表不存在")
                continue
            tws = twb[s]
            thdr = header_map_of(tws)
            if tgt_key_col not in thdr:
                log(f"[跳过] {label}: 表头里找不到键列「{tgt_key_col}」")
                continue
            tgt_fields = {}
            for g in src_fields:
                col = groups.get(g, {}).get(tgt)
                if not col:
                    log(f"[提示] {label}: 在映射组「{g}」里没有登记字段，该组跳过")
                    continue
                if col not in thdr:
                    log(f"[提示] {label}: 表头里找不到字段列「{col}」，该组跳过")
                    continue
                tgt_fields[g] = thdr[col]
            if not tgt_fields:
                log(f"[跳过] {label}: 没有可同步的字段")
                continue

            # 生成变更计划（不立即修改）
            kc = thdr[tgt_key_col]
            changes = []
            hit_keys = set()
            formula_tgt = 0  # 目标是公式而被拦截的单元格数
            for r in range(2, tws.max_row + 1):
                k = norm(tws.cell(r, kc).value)
                if not k or k not in src_data:
                    continue
                hit_keys.add(k)
                for g, c in tgt_fields.items():
                    v = src_data[k][g]
                    if v is None or norm(v) == "":
                        continue  # 源值为空不同步
                    cell = tws.cell(r, c)
                    if is_formula_like(cell.value):
                        formula_tgt += 1  # 目标是公式，不覆盖
                        continue
                    if norm(cell.value) != norm(v):
                        changes.append((r, g, cell.value, v))
            if changes:
                total_changes += len(changes)
                file_changes += len(changes)
                log(f"\n{'[预览] ' if dry else ''}{label}: {len(changes)} 处改动")
                for r, g, old, new in changes[:200]:
                    log(f"  行{r} 「{groups[g][tgt]}」: {old!r} -> {new!r}")
                if len(changes) > 200:
                    log(f"  …… 其余 {len(changes) - 200} 处省略")
                if not dry:
                    for r, g, old, new in changes:
                        tws.cell(r, tgt_fields[g]).value = new
            else:
                log(f"{label}: 无需改动（值已一致）")
            if formula_tgt:
                tot_blocked += formula_tgt
                log(f"  [警告] {label}: {formula_tgt} 个目标单元格是公式，未覆盖（工具不改动公式）")
            missing = set(src_data) - hit_keys
            if missing:
                tot_unmatched += len(missing)
                sample = ", ".join(sorted(missing)[:10])
                log(f"  [提示] {label}: 有 {len(missing)} 个数据源的键值未匹配到行（{sample}{' ……' if len(missing) > 10 else ''}）")

        if not file_changes:
            results.append(f"{f}: 无需改动")
        elif dry:
            results.append(f"{f}: 预览 {file_changes} 处改动")
        else:
            try:
                bak = save_with_backup(twb, f, base_dir, backup_dir, stamp)
                log(f"[已保存] {f}（本批共 {file_changes} 处改动，已按原始状态备份）")
                results.append(f"{f}: {file_changes} 处改动，已保存"
                               + (f"（备份: {bak}）" if bak else "（未备份）"))
            except Exception as ex:
                log(f"[错误] {f}: 保存失败: {ex}（原文件未被修改）")
                results.append(f"{f}: 保存失败（{ex}），原文件未被修改")
        twb.close()

    tot_blocked += len(formula_src)
    out = [f"【同步】{datetime.now():%Y-%m-%d %H:%M:%S}"]
    out.append(f"配置: 键={key_group} | 源={src[0]}|{src[1]} | 同步组={', '.join(sync_groups)} | 目标={len(by_file)} 个文件")
    out.append("─" * 46)
    out.append("明细:")
    out.extend(f"  {l}" for l in _d)
    out.append("警告:")
    if _w:
        out.extend(f"  {l}" for l in _w)
    else:
        out.append("  无")
    out.append("─" * 46)
    out.append("执行结果:")
    out.extend(f"  {l}" for l in results)
    out.append(f"汇总: 改动 {total_changes} | 公式拦截 {tot_blocked} | 未匹配 {tot_unmatched} | 状态: {'预览(未写入)' if dry else '成功'}")
    _log_out("\n".join(out))
    return total_changes


def run_replace(base_dir, groups, src_group, dst_group, ref, targets, backup_dir, dry, log):
    """字段替换: 把 targets 各列中属于 src_group 的值，按参照表 ref 翻译成 dst_group 的值。
    ref = (文件, 表格); targets = [(文件, 表格, 字段名), ...]。
    只改数据区（第 2 行起），表头不动。返回替换单元格数。
    日志按统一模板输出: 标题+配置 → 明细 → 警告 → 执行结果 → 汇总。"""
    _d, _w = [], []
    _log_out = log

    def log(msg):
        (_w if msg.strip().startswith("[警告]") else _d).append(msg)

    def flush():
        if _d or _w:
            _log_out("\n".join(_d + _w))
            _d.clear()
            _w.clear()

    tot_blocked = 0    # 公式拦截数
    tot_unmatched = 0  # 参照表查不到的值个数

    src_col = groups.get(src_group, {}).get(ref)
    dst_col = groups.get(dst_group, {}).get(ref)
    if not src_col or not dst_col:
        log("[错误] 参照表没有同时登记源组和目标组的字段")
        flush()
        return 0
    try:
        rwb, rws, rhdr = read_sheet(base_dir / ref[0], ref[1])
    except Exception as ex:
        log(f"[错误] 读取参照表失败: {ex}")
        flush()
        return 0
    if src_col not in rhdr or dst_col not in rhdr:
        log(f"[错误] 参照表表头里找不到列「{src_col}」或「{dst_col}」")
        rwb.close()
        flush()
        return 0
    mapping = {}
    dup = 0
    formula_ref = 0  # 参照值是公式的条数
    for r in range(2, rws.max_row + 1):
        k = norm(rws.cell(r, rhdr[src_col]).value)
        if not k:
            continue
        if k in mapping:
            dup += 1
            continue
        v_dst = rws.cell(r, rhdr[dst_col]).value
        if is_formula_like(v_dst):
            formula_ref += 1
            v_dst = None  # 公式不作为参照值，视同空值跳过
        mapping[k] = v_dst
    dst_vals = {norm(v) for v in mapping.values() if norm(v)}
    rwb.close()
    if dup:
        log(f"[提示] 参照表有 {dup} 个重复的「{src_group}」值，已取首次出现的行")
    if formula_ref:
        log(f"[警告] 参照表有 {formula_ref} 条对照的取值是公式，已跳过（工具不处理公式内容）")
    log(f"参照表: {ref[0]} | {ref[1]}（{src_col} → {dst_col}），共 {len(mapping)} 条对照")
    log(f"翻译规则: 「{src_group}」→「{dst_group}」")

    total = 0
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    # 按文件分批: 每个文件 打开一次 → 校验并生成变更计划 → 备份一次 → 应用全部修改 → 原子保存
    by_file = {}
    for f, s, col in targets:
        by_file.setdefault(f, []).append((s, col))

    results = []
    for f, entries in by_file.items():
        try:
            twb = openpyxl.load_workbook(base_dir / f)
        except Exception as ex:
            log(f"[错误] {f}: 打开失败: {ex}")
            results.append(f"{f}: 打开失败（{ex}）")
            continue
        file_changes = 0
        for s, col in entries:
            label = f"{f} | {s} | {col}"
            if s not in twb.sheetnames:
                log(f"[错误] {label}: 工作表不存在")
                continue
            tws = twb[s]
            thdr = header_map_of(tws)
            if col not in thdr:
                log(f"[跳过] {label}: 表头里找不到该列")
                continue
            # 生成变更计划（不立即修改）
            c = thdr[col]
            changes = []
            unmatched = {}
            already = 0
            empty_dst = 0
            formula_tgt = 0  # 目标是公式而被拦截的单元格数
            for r in range(2, tws.max_row + 1):
                cell = tws.cell(r, c)
                v = cell.value
                if v is None or norm(v) == "":
                    continue
                if is_formula_like(v):
                    formula_tgt += 1  # 目标是公式，不替换
                    continue
                k = norm(v)
                if k in mapping:
                    pass  # 待替换
                elif k in dst_vals:
                    already += 1  # 已是目标值，跳过
                    continue
                else:
                    unmatched[k] = unmatched.get(k, 0) + 1
                    continue
                new = mapping[k]
                if new is None or norm(new) == "":
                    empty_dst += 1
                    continue
                if norm(v) != norm(new):
                    changes.append((r, v, new))
            if changes:
                total += len(changes)
                file_changes += len(changes)
                log(f"\n{'[预览] ' if dry else ''}{label}: {len(changes)} 处替换")
                for r, old, new in changes[:200]:
                    log(f"  行{r}: {old!r} -> {new!r}")
                if len(changes) > 200:
                    log(f"  …… 其余 {len(changes) - 200} 处省略")
                if not dry:
                    for r, old, new in changes:
                        tws.cell(r, c).value = new
            else:
                log(f"{label}: 无需替换")
            if unmatched:
                tot_unmatched += sum(unmatched.values())
                items = ", ".join(f"{k!r}×{n}" for k, n in sorted(unmatched.items())[:10])
                log(f"  [提示] {label}: {sum(unmatched.values())} 个值在参照表找不到（保留原值）: {items}"
                    + (" ……" if len(unmatched) > 10 else ""))
            if already:
                log(f"  [提示] {label}: {already} 个单元格已是「{dst_group}」的值，已跳过")
            if formula_tgt:
                tot_blocked += formula_tgt
                log(f"  [警告] {label}: {formula_tgt} 个目标单元格是公式，未替换（工具不改动公式）")
            if empty_dst:
                log(f"  [提示] {label}: {empty_dst} 个值在参照表中的目标列为空（保留原值）")

        if not file_changes:
            results.append(f"{f}: 无需改动")
        elif dry:
            results.append(f"{f}: 预览 {file_changes} 处替换")
        else:
            try:
                bak = save_with_backup(twb, f, base_dir, backup_dir, stamp)
                log(f"[已保存] {f}（本批共 {file_changes} 处替换，已按原始状态备份）")
                results.append(f"{f}: {file_changes} 处替换，已保存"
                               + (f"（备份: {bak}）" if bak else "（未备份）"))
            except Exception as ex:
                log(f"[错误] {f}: 保存失败: {ex}（原文件未被修改）")
                results.append(f"{f}: 保存失败（{ex}），原文件未被修改")
        twb.close()

    tot_blocked += formula_ref
    out = [f"【替换-按参照表】{datetime.now():%Y-%m-%d %H:%M:%S}"]
    out.append(f"配置: 规则={src_group}→{dst_group} | 参照={ref[0]}|{ref[1]} | 目标={len(by_file)} 个文件")
    out.append("─" * 46)
    out.append("明细:")
    out.extend(f"  {l}" for l in _d)
    out.append("警告:")
    if _w:
        out.extend(f"  {l}" for l in _w)
    else:
        out.append("  无")
    out.append("─" * 46)
    out.append("执行结果:")
    out.extend(f"  {l}" for l in results)
    out.append(f"汇总: 替换 {total} | 公式拦截 {tot_blocked} | 未匹配 {tot_unmatched} | 状态: {'预览(未写入)' if dry else '成功'}")
    _log_out("\n".join(out))
    return total


def run_replace_align(base_dir, ref, match_col, value_col, targets, backup_dir, dry, log):
    """按字段匹配替换: 目标列不按自身值查, 而是取同一行的匹配字段值,
    去参照表(ref)的匹配列(match_col)对齐行, 把该行的取值列(value_col)写入目标列。
    目标子表需有与 match_col 同名的列。只改数据区（第 2 行起），表头不动。
    ref = (文件, 表格); targets = [(文件, 表格, 字段名), ...]。返回替换单元格数。
    日志按统一模板输出: 标题+配置 → 明细 → 警告 → 执行结果 → 汇总。"""
    _d, _w = [], []
    _log_out = log

    def log(msg):
        (_w if msg.strip().startswith("[警告]") else _d).append(msg)

    def flush():
        if _d or _w:
            _log_out("\n".join(_d + _w))
            _d.clear()
            _w.clear()

    tot_blocked = 0
    tot_unmatched = 0

    try:
        rwb, rws, rhdr = read_sheet(base_dir / ref[0], ref[1])
    except Exception as ex:
        log(f"[错误] 读取参照表失败: {ex}")
        flush()
        return 0
    if match_col not in rhdr or value_col not in rhdr:
        log(f"[错误] 参照表表头里找不到列「{match_col}」或「{value_col}」")
        rwb.close()
        flush()
        return 0
    mapping = {}
    dup = 0
    formula_ref = 0  # 参照值是公式的条数
    for r in range(2, rws.max_row + 1):
        k = norm(rws.cell(r, rhdr[match_col]).value)
        if not k:
            continue
        if k in mapping:
            dup += 1
            continue
        v_dst = rws.cell(r, rhdr[value_col]).value
        if is_formula_like(v_dst):
            formula_ref += 1
            v_dst = None  # 公式不作为参照值，视同空值跳过
        mapping[k] = v_dst
    dst_vals = {norm(v) for v in mapping.values() if norm(v)}
    rwb.close()
    if dup:
        log(f"[提示] 参照表有 {dup} 个重复的「{match_col}」值，已取首次出现的行")
    if formula_ref:
        log(f"[警告] 参照表有 {formula_ref} 条对照的取值是公式，已跳过（工具不处理公式内容）")
    log(f"参照表: {ref[0]} | {ref[1]}（按「{match_col}」对齐，取「{value_col}」），共 {len(mapping)} 条对照")

    total = 0
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    by_file = {}
    for f, s, col in targets:
        by_file.setdefault(f, []).append((s, col))

    results = []
    for f, entries in by_file.items():
        try:
            twb = openpyxl.load_workbook(base_dir / f)
        except Exception as ex:
            log(f"[错误] {f}: 打开失败: {ex}")
            results.append(f"{f}: 打开失败（{ex}）")
            continue
        file_changes = 0
        for s, col in entries:
            label = f"{f} | {s} | {col}"
            if s not in twb.sheetnames:
                log(f"[错误] {label}: 工作表不存在")
                continue
            tws = twb[s]
            thdr = header_map_of(tws)
            if col not in thdr:
                log(f"[跳过] {label}: 表头里找不到该列")
                continue
            if match_col not in thdr:
                log(f"[跳过] {label}: 该表没有匹配列「{match_col}」，无法对齐")
                continue
            if col == match_col:
                log(f"[跳过] {label}: 目标列与匹配列相同")
                continue
            mc, c = thdr[match_col], thdr[col]
            changes = []
            unmatched = {}
            already = 0
            empty_dst = 0
            formula_tgt = 0  # 目标是公式而被拦截的单元格数
            for r in range(2, tws.max_row + 1):
                kv = norm(tws.cell(r, mc).value)
                if not kv:
                    continue
                if kv not in mapping:
                    unmatched[kv] = unmatched.get(kv, 0) + 1
                    continue
                new = mapping[kv]
                if new is None or norm(new) == "":
                    empty_dst += 1
                    continue
                cell = tws.cell(r, c)
                if is_formula_like(cell.value):
                    formula_tgt += 1  # 目标是公式，不替换
                    continue
                if norm(cell.value) == norm(new):
                    already += 1
                    continue
                changes.append((r, cell.value, new))
            if changes:
                total += len(changes)
                file_changes += len(changes)
                log(f"\n{'[预览] ' if dry else ''}{label}: {len(changes)} 处替换")
                for r, old, new in changes[:200]:
                    log(f"  行{r}: {old!r} -> {new!r}")
                if len(changes) > 200:
                    log(f"  …… 其余 {len(changes) - 200} 处省略")
                if not dry:
                    for r, old, new in changes:
                        tws.cell(r, c).value = new
            else:
                log(f"{label}: 无需替换")
            if unmatched:
                tot_unmatched += sum(unmatched.values())
                items = ", ".join(f"{k!r}×{n}" for k, n in sorted(unmatched.items())[:10])
                log(f"  [提示] {label}: {sum(unmatched.values())} 行的「{match_col}」在参照表找不到（保留原值）: {items}"
                    + (" ……" if len(unmatched) > 10 else ""))
            if already:
                log(f"  [提示] {label}: {already} 个单元格已与参照值一致，已跳过")
            if formula_tgt:
                tot_blocked += formula_tgt
                log(f"  [警告] {label}: {formula_tgt} 个目标单元格是公式，未替换（工具不改动公式）")
            if empty_dst:
                log(f"  [提示] {label}: {empty_dst} 行在参照表中的「{value_col}」为空（保留原值）")

        if not file_changes:
            results.append(f"{f}: 无需改动")
        elif dry:
            results.append(f"{f}: 预览 {file_changes} 处替换")
        else:
            try:
                bak = save_with_backup(twb, f, base_dir, backup_dir, stamp)
                log(f"[已保存] {f}（本批共 {file_changes} 处替换，已按原始状态备份）")
                results.append(f"{f}: {file_changes} 处替换，已保存"
                               + (f"（备份: {bak}）" if bak else "（未备份）"))
            except Exception as ex:
                log(f"[错误] {f}: 保存失败: {ex}（原文件未被修改）")
                results.append(f"{f}: 保存失败（{ex}），原文件未被修改")
        twb.close()

    tot_blocked += formula_ref
    out = [f"【替换-按字段匹配】{datetime.now():%Y-%m-%d %H:%M:%S}"]
    out.append(f"配置: 参照={ref[0]}|{ref[1]} | 对齐={match_col} | 取值={value_col} | 目标={len(by_file)} 个文件")
    out.append("─" * 46)
    out.append("明细:")
    out.extend(f"  {l}" for l in _d)
    out.append("警告:")
    if _w:
        out.extend(f"  {l}" for l in _w)
    else:
        out.append("  无")
    out.append("─" * 46)
    out.append("执行结果:")
    out.extend(f"  {l}" for l in results)
    out.append(f"汇总: 替换 {total} | 公式拦截 {tot_blocked} | 未匹配 {tot_unmatched} | 状态: {'预览(未写入)' if dry else '成功'}")
    _log_out("\n".join(out))
    return total


class AddDialog(tk.Toplevel):
    """登记字段窗口：文件 → 子表 → 未登记字段（多选）→ 映射组（可选/新建）→ 加入 → 保存。"""

    def __init__(self, master, on_saved):
        super().__init__(master)
        self.title("登记字段")
        self.geometry("1080x520")
        self.on_saved = on_saved
        self.base_dir = master.base_dir
        self.staged = []  # [(文件, 表格, 字段名, 映射组)]
        self._files = []    # [(文件, [(子表, [表头...]), ...])]
        self._cur_hdr = []

        paned = ttk.Frame(self, padding=8)
        paned.pack(fill=tk.BOTH, expand=True)

        f1 = ttk.LabelFrame(paned, text="1. 文件", padding=4)
        f1.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.file_tree = ttk.Treeview(f1, show="tree", selectmode="browse", height=16)
        sb1 = ttk.Scrollbar(f1, orient="vertical", command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=sb1.set)
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb1.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_tree.bind("<<TreeviewSelect>>", lambda e: self._on_file_select())

        f2 = ttk.LabelFrame(paned, text="2. 子表", padding=4)
        f2.grid(row=0, column=1, sticky="nsew", padx=4)
        self.sheet_list = tk.Listbox(f2, width=20, selectmode=tk.BROWSE, exportselection=False)
        self.sheet_list.pack(fill=tk.BOTH, expand=True)
        self.sheet_list.bind("<<ListboxSelect>>", lambda e: self._on_sheet_select())

        f3 = ttk.LabelFrame(paned, text="3. 未登记字段（多选）", padding=4)
        f3.grid(row=0, column=2, sticky="nsew", padx=4)
        self.dlg_field_filter = tk.StringVar()
        self.dlg_field_filter.trace_add("write", lambda *a: self._apply_dlg_field_filter())
        fe = ttk.Entry(f3, textvariable=self.dlg_field_filter)
        fe.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(f3, text="↑ 输入关键字筛选字段", foreground="#888888").pack(anchor=tk.W)
        self.field_list = tk.Listbox(f3, selectmode=tk.EXTENDED, width=18, exportselection=False)
        self.field_list.pack(fill=tk.BOTH, expand=True)

        f4 = ttk.Frame(paned, padding=4)
        f4.grid(row=0, column=3, sticky="nsew", padx=(4, 0))
        ttk.Label(f4, text="4. 映射组（可选已有\n或直接输入新组名）:").pack(anchor=tk.W)
        self.group_var = tk.StringVar()
        self.group_box = ttk.Combobox(f4, textvariable=self.group_var, width=18)
        self.group_box.pack(anchor=tk.W, pady=4)
        ttk.Button(f4, text="加入 →", command=self._stage).pack(anchor=tk.W, pady=4)

        f5 = ttk.LabelFrame(paned, text="待保存的登记", padding=4)
        f5.grid(row=0, column=4, sticky="nsew", padx=(4, 0))
        self.staged_list = tk.Listbox(f5, width=30, exportselection=False)
        self.staged_list.pack(fill=tk.BOTH, expand=True)
        ttk.Button(f5, text="移除选中", command=self._unstage).pack(anchor=tk.E, pady=2)

        for c in range(5):
            paned.columnconfigure(c, weight=1)
        paned.rowconfigure(0, weight=1)

        bar = ttk.Frame(self, padding=8)
        bar.pack(fill=tk.X)
        self.status = ttk.Label(bar, text="")
        self.status.pack(side=tk.LEFT)
        ttk.Button(bar, text="保存并关闭", command=self._save).pack(side=tk.RIGHT, padx=4)
        ttk.Button(bar, text="刷新列表", command=self.refresh).pack(side=tk.RIGHT)

        self.refresh()
        self.transient(master)
        self.grab_set()

    def refresh(self):
        try:
            tables, groups = load_tool(self.master.tool_path)
        except Exception as ex:
            messagebox.showerror("读取工具表失败", str(ex), parent=self)
            return
        self._groups = groups
        self.group_box["values"] = sorted(groups)
        # 每个表格已登记的字段
        self._reg_fields = {}
        for m in groups.values():
            for (f, s), col in m.items():
                self._reg_fields.setdefault((f, s), set()).add(col)
        self._files = scan_tables(self.base_dir, self.master.tool_path, self.master.backup_dir,
                                  cache_dir=self.master.app_dir)
        self._file_map = dict(self._files)
        self.file_tree.delete(*self.file_tree.get_children())
        build_file_tree(self.file_tree, [f for f, _ in self._files])
        self.sheet_list.delete(0, tk.END)
        self._cur_hdr = []
        self._apply_dlg_field_filter()
        self.status.config(text=f"共 {len(self._files)} 个文件")

    def _on_file_select(self):
        self.sheet_list.delete(0, tk.END)
        self._cur_hdr = []
        self._apply_dlg_field_filter()
        rel = tree_selected_file(self.file_tree)
        if rel is None:
            return
        for s, _ in self._file_map[rel]:
            self.sheet_list.insert(tk.END, s)

    def _on_sheet_select(self):
        f = tree_selected_file(self.file_tree)
        si = self.sheet_list.curselection()
        self._cur_hdr = []
        if f and si:
            s, hdr = self._file_map[f][si[0]]
            registered = self._reg_fields.get((f, s), set())
            self._cur_hdr = [h for h in hdr if h not in registered]
            if not self._cur_hdr:
                self.status.config(text=f"「{s}」的字段已全部登记")
            else:
                self.status.config(text=f"「{s}」还有 {len(self._cur_hdr)} 个未登记字段")
        self._apply_dlg_field_filter()

    def _apply_dlg_field_filter(self):
        kw = self.dlg_field_filter.get().strip().lower()
        self.field_list.delete(0, tk.END)
        for h in getattr(self, "_cur_hdr", []):
            if not kw or kw in h.lower():
                self.field_list.insert(tk.END, h)

    def _stage(self):
        f = tree_selected_file(self.file_tree)
        si = self.sheet_list.curselection()
        g = self.group_var.get().strip()
        if not f or not si:
            messagebox.showwarning("缺少选择", "请先选择文件和子表", parent=self)
            return
        if not g:
            messagebox.showwarning("缺少选择", "请选择或输入映射组", parent=self)
            return
        fsel = self.field_list.curselection()
        if not fsel:
            messagebox.showwarning("缺少选择", "请勾选要加入的字段", parent=self)
            return
        s = self._file_map[f][si[0]][0]
        for i in fsel:
            col = self.field_list.get(i)
            if col not in self._cur_hdr:
                continue
            item = (f, s, col, g)
            # 同一表格在同一映射组里只保留一个字段：替换待保存列表中的旧条目
            for j in range(len(self.staged) - 1, -1, -1):
                sf, ss, sc, sg = self.staged[j]
                if (sf, ss, sg) == (f, s, g) and sc != col:
                    del self.staged[j]
                    self.staged_list.delete(j)
            if item not in self.staged:
                self.staged.append(item)
                self.staged_list.insert(tk.END, f"{g} | {f} | {s} | {col}")
        self.status.config(text=f"待保存 {len(self.staged)} 条")

    def _unstage(self):
        for i in reversed(self.staged_list.curselection()):
            self.staged_list.delete(i)
            del self.staged[i]
        self.status.config(text=f"待保存 {len(self.staged)} 条")

    def _save(self):
        if not self.staged:
            messagebox.showwarning("没有内容", "没有待保存的登记", parent=self)
            return
        added, updated = add_associations(self.master.tool_path, self.staged)
        messagebox.showinfo("保存完成", f"已新增 {added} 条字段登记，更新 {updated} 条", parent=self)
        self.on_saved()
        self.destroy()


class DeleteGroupDialog(tk.Toplevel):
    """删除映射组: 选组 → 确认 → 删除该组的所有字段登记。"""

    def __init__(self, master, on_done):
        super().__init__(master)
        self.title("删除映射组")
        self.resizable(False, False)
        self.on_done = on_done
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="选择要删除的映射组（只删映射关系，不动数据表）:").pack(anchor=tk.W)
        self.grp_var = tk.StringVar()
        self.grp_box = ttk.Combobox(frm, textvariable=self.grp_var, state="readonly", width=34)
        self.grp_box.pack(anchor=tk.W, pady=6)
        bar = ttk.Frame(frm)
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="删除该组", command=self._delete).pack(side=tk.LEFT)
        ttk.Button(bar, text="关闭", command=self.destroy).pack(side=tk.RIGHT)
        self._refresh()
        self.transient(master)
        self.grab_set()

    def _refresh(self):
        self._groups = load_tool(self.master.tool_path)[1]
        self.grp_box["values"] = [f"{g}（{len(m)} 条）" for g, m in sorted(self._groups.items())]
        self.grp_var.set("")

    def _delete(self):
        pos = self.grp_box.current()
        if pos < 0:
            messagebox.showwarning("缺少选择", "请先选择映射组", parent=self)
            return
        g = sorted(self._groups)[pos]
        n = len(self._groups[g])
        if not messagebox.askyesno(
            "确认删除",
            f"将删除映射组「{g}」的全部 {n} 条字段登记。\n（只删除工具表里的映射，不会改动任何数据表）\n\n确定？",
            parent=self):
            return
        deleted = delete_group(self.master.tool_path, g)
        messagebox.showinfo("已删除", f"已删除映射组「{g}」的 {deleted} 条登记", parent=self)
        self.on_done()
        self._refresh()


class SetupDialog(tk.Toplevel):
    """工具表设置引导: 选择已有表格作为工具表, 或新建一个。
    current 不为 None 时表示「更换」场景（已有在用的工具表）。"""

    def __init__(self, master, current=None):
        super().__init__(master)
        self.result = None
        frm = ttk.Frame(self, padding=20)
        frm.pack(fill=tk.BOTH, expand=True)
        if current:
            self.title("更换工具表")
            text = (f"当前工具表：{current}\n\n"
                    "你可以选择其他表格作为工具表，或新建一个。\n"
                    "点「取消」则保持当前工具表不变。")
            cancel_text = "取消"
        else:
            self.title("设置工具表")
            text = ("还没有设置工具表。\n\n"
                    "工具表用来登记表格和字段映射，\n"
                    "你可以选择一个已有的 xlsx 表格作为工具表，或新建一个。")
            cancel_text = "退出"
        self.resizable(False, False)
        ttk.Label(frm, text=text, justify=tk.LEFT).pack(pady=(0, 16))
        ttk.Button(frm, text="选择已有表格作为工具表",
                   command=lambda: self._done("pick")).pack(fill=tk.X, pady=3)
        ttk.Button(frm, text="新建一个工具表",
                   command=lambda: self._done("new")).pack(fill=tk.X, pady=3)
        ttk.Button(frm, text=cancel_text,
                   command=lambda: self._done(None)).pack(fill=tk.X, pady=(12, 0))
        self.protocol("WM_DELETE_WINDOW", lambda: self._done(None))
        self.transient(master)
        self.grab_set()
        self.update_idletasks()
        # 居中到主窗口
        x = master.winfo_x() + (master.winfo_width() - self.winfo_width()) // 2
        y = master.winfo_y() + (master.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _done(self, result):
        self.result = result
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("字段映射同步工具")
        self.geometry("1020x680")
        self.app_dir = app_dir()          # 配置文件所在目录（跟随 exe/脚本）
        self.base_dir = self.app_dir      # 数据目录（跟随工具表，选定后更新）
        self.tool_path = self._ensure_tool_file()
        if not self.tool_path:
            self.destroy()
            return
        self.base_dir = self.tool_path.resolve().parent
        self.backup_dir = find_backup_config(self.app_dir) or (self.base_dir / BACKUP_DIR)
        self.title(f"字段映射同步工具 - {self.tool_path.name}")
        if not self._reload():
            self.destroy()
            return
        self._build()
        self._startup_report()

    def _set_backup_dir(self):
        d = filedialog.askdirectory(parent=self, initialdir=str(self.backup_dir),
                                    title="选择备份目录")
        if not d:
            return
        self.backup_dir = Path(d)
        save_backup_config(self.app_dir, self.backup_dir)
        self._log(f"[备份目录] 已设置为: {self.backup_dir}")

    def _startup_report(self):
        """在日志区输出启动校验结果，并检查失效映射。"""
        total = sum(len(m) for m in self.groups.values())
        self._log(f"[启动校验] 工具表「{self.tool_path.name}」校验通过（数据目录: {self.base_dir}）")
        self._log(f"[启动校验] 备份目录: {self.backup_dir}")
        self._log(f"[启动校验] 登记表格 {len(self.tables)} 个，映射组 {len(self.groups)} 个")
        lost = check_mappings(self.base_dir, self.tables, self.groups,
                              cache_dir=self.app_dir, tool_path=self.tool_path,
                              exclude_dir=self.backup_dir)
        ok = total - len(lost)
        self._log(f"[启动校验] 字段映射校验完成：共 {total} 条，正常 {ok} 条，失效 {len(lost)} 条")
        self._log("-" * 60)
        self._handle_lost_mappings(lost)

    def _ensure_tool_file(self, force=False):
        """确定工具表: 先用上次记录(并校验格式), 没有则引导用户选择或新建。
        force=True 时跳过记录直接进入引导（用于「更换工具表」）。"""
        if not force:
            p = find_tool_config(self.app_dir)
            if p:
                if check_tool_format(p) == "ok":
                    return p
                messagebox.showwarning(
                    "工具表不可用",
                    f"上次使用的工具表「{p.name}」已不符合工具表格式，请重新设置。",
                    parent=self)
            else:
                # 没有配置记录时, 同目录下默认文件名的合规表格直接作为工具表
                default = self.app_dir / TOOL_FILE
                if default.exists() and check_tool_format(default) == "ok":
                    save_tool_config(self.app_dir, default)
                    return default
        while True:
            dlg = SetupDialog(self, current=getattr(self, "tool_path", None) if force else None)
            self.wait_window(dlg)
            if dlg.result is None:
                return None
            if dlg.result == "pick":
                f = filedialog.askopenfilename(
                    parent=self, initialdir=self.app_dir, title="选择作为工具表的 xlsx",
                    filetypes=[("Excel 表格", "*.xlsx")])
                if not f:
                    continue
                status = check_tool_format(f)
                if status == "ok":
                    save_tool_config(self.app_dir, Path(f))
                    return Path(f)
                if status == "empty":
                    init_tool_file(f)
                    save_tool_config(self.app_dir, Path(f))
                    messagebox.showinfo("初始化完成",
                                        "已在该表格中写入「使用说明」「总表」「字段表」，可以开始使用了。", parent=self)
                    return Path(f)
                messagebox.showerror(
                    "无法使用该表格",
                    "该表格不符合工具表格式，也不是空表，无法作为工具表。\n请换一个表格，或选择新建一个工具表。",
                    parent=self)
                continue
            if dlg.result == "new":
                f = filedialog.asksaveasfilename(
                    parent=self, initialdir=self.app_dir, initialfile=TOOL_FILE,
                    defaultextension=".xlsx", title="新建工具表",
                    filetypes=[("Excel 表格", "*.xlsx")])
                if not f:
                    continue
                if Path(f).exists() and check_tool_format(f) == "invalid":
                    messagebox.showerror(
                        "无法使用该表格",
                        "该文件已存在且不符合工具表格式、也不是空表。\n请换一个文件名，或选择已有表格。",
                        parent=self)
                    continue
                init_tool_file(f)
                save_tool_config(self.app_dir, Path(f))
                return Path(f)

    def _reload(self):
        try:
            self.tables, self.groups = load_tool(self.tool_path)
            return True
        except Exception as ex:
            messagebox.showerror("读取工具表失败", str(ex))
            return False

    def _build(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=False, padx=4, pady=(4, 0))

        # ---------- 页签1: 同步 ----------
        tab_sync = ttk.Frame(self.nb)
        self.nb.add(tab_sync, text="  同步  ")

        top = ttk.Frame(tab_sync, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text="① 同步映射组:").grid(row=0, column=0, sticky=tk.W)
        self.grp_var = tk.StringVar()
        self.grp_box = ttk.Combobox(top, textvariable=self.grp_var, state="readonly", width=18)
        self.grp_box.grid(row=0, column=1, sticky=tk.W, padx=(4, 16))
        self.grp_box.bind("<<ComboboxSelected>>", lambda e: self._on_grp_change())

        ttk.Label(top, text="② 同步源字段（以此数值为准）:").grid(row=0, column=2, sticky=tk.W)
        self.src_var = tk.StringVar()
        self.src_box = ttk.Combobox(top, textvariable=self.src_var, state="readonly", width=52)
        self.src_box.grid(row=0, column=3, sticky=tk.W, padx=4)
        self.src_box.bind("<<ComboboxSelected>>", lambda e: self._on_src_change())

        mid = ttk.Frame(tab_sync, padding=8)
        mid.pack(fill=tk.BOTH, expand=False)
        f2 = ttk.LabelFrame(mid, text="③ 被同步字段（多选，格式: 文件 | 表格 | 字段名）", padding=4)
        f2.pack(fill=tk.BOTH, expand=True)
        self.tgt_list = tk.Listbox(f2, selectmode=tk.EXTENDED, height=8, exportselection=False)
        self.tgt_list.pack(fill=tk.BOTH, expand=True)
        self.tgt_list.bind("<<ListboxSelect>>", lambda e: self._refresh_key_candidates())
        ttk.Button(f2, text="全选", command=self._select_all_targets).pack(anchor=tk.E, pady=2)

        keyf = ttk.Frame(tab_sync, padding=(8, 0, 8, 4))
        keyf.pack(fill=tk.X)
        ttk.Label(keyf, text="④ 键映射组（用于对齐行，已按所选字段筛选）:").pack(side=tk.LEFT)
        self.key_var = tk.StringVar()
        self.key_box = ttk.Combobox(keyf, textvariable=self.key_var, state="readonly", width=24)
        self.key_box.pack(side=tk.LEFT, padx=4)

        # ---------- 页签2: 字段替换 ----------
        tab_rp = ttk.Frame(self.nb)
        self.nb.add(tab_rp, text="  字段替换  ")

        rmid = ttk.LabelFrame(tab_rp, text="① 被替换字段（逐级: 文件 → 子表 → 字段，可多选；输入关键字则按完整路径筛选）", padding=4)
        rmid.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 0))
        fr = ttk.Frame(rmid)
        fr.pack(fill=tk.X, pady=(0, 2))
        self.rp_filter = tk.StringVar()
        self.rp_filter.trace_add("write", lambda *a: self._apply_rp_filter())
        ttk.Entry(fr, textvariable=self.rp_filter).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.rp_count = ttk.Label(fr, text="已选 0 个字段")
        self.rp_count.pack(side=tk.LEFT, padx=6)

        pane = ttk.Frame(rmid)
        pane.pack(fill=tk.BOTH, expand=True)
        pane.columnconfigure(0, weight=0)
        pane.columnconfigure(1, weight=0)
        pane.columnconfigure(2, weight=1)
        pf1 = ttk.Frame(pane)
        pf1.grid(row=0, column=0, sticky="nsew", padx=(0, 3))
        ttk.Label(pf1, text="文件").pack(anchor=tk.W)
        self.rp_file_tree = ttk.Treeview(pf1, show="tree", selectmode="browse", height=8)
        self.rp_file_tree.pack(fill=tk.BOTH, expand=True)
        self.rp_file_tree.bind("<<TreeviewSelect>>", lambda e: self._rp_on_file())
        pf2 = ttk.Frame(pane)
        pf2.grid(row=0, column=1, sticky="nsew", padx=3)
        ttk.Label(pf2, text="子表").pack(anchor=tk.W)
        self.rp_sheet_list = tk.Listbox(pf2, selectmode=tk.BROWSE, width=14, height=8, exportselection=False)
        self.rp_sheet_list.pack(fill=tk.BOTH, expand=True)
        self.rp_sheet_list.bind("<<ListboxSelect>>", lambda e: self._rp_on_sheet())
        pf3 = ttk.Frame(pane)
        pf3.grid(row=0, column=2, sticky="nsew", padx=(3, 0))
        ttk.Label(pf3, text="字段（多选）").pack(anchor=tk.W)
        self.rp_list = tk.Listbox(pf3, selectmode=tk.EXTENDED, width=46, height=8, exportselection=False)
        self.rp_list.pack(fill=tk.BOTH, expand=True)
        self.rp_list.bind("<<ListboxSelect>>", lambda e: self._rp_on_select())

        rbtn = ttk.Frame(rmid)
        rbtn.pack(fill=tk.X, pady=2)
        ttk.Button(rbtn, text="全选当前列表", command=self._rp_select_all_shown).pack(side=tk.LEFT)
        ttk.Button(rbtn, text="清空已选", command=self._rp_clear_selected).pack(side=tk.LEFT, padx=6)

        self.nb2 = ttk.Notebook(tab_rp)
        self.nb2.pack(fill=tk.X, padx=4, pady=4)

        rtop = ttk.Frame(self.nb2, padding=8)
        self.nb2.add(rtop, text="按参照表（映射组）")
        ttk.Label(rtop, text="② 当前值是:").grid(row=0, column=0, sticky=tk.W)
        self.rp_src_var = tk.StringVar()
        self.rp_src_box = ttk.Combobox(rtop, textvariable=self.rp_src_var, state="readonly", width=16)
        self.rp_src_box.grid(row=0, column=1, sticky=tk.W, padx=(4, 12))
        self.rp_src_box.bind("<<ComboboxSelected>>", lambda e: self._refresh_rp_ref())
        ttk.Label(rtop, text="③ 替换成:").grid(row=0, column=2, sticky=tk.W)
        self.rp_dst_var = tk.StringVar()
        self.rp_dst_box = ttk.Combobox(rtop, textvariable=self.rp_dst_var, state="readonly", width=16)
        self.rp_dst_box.grid(row=0, column=3, sticky=tk.W, padx=(4, 12))
        self.rp_dst_box.bind("<<ComboboxSelected>>", lambda e: self._refresh_rp_ref())
        ttk.Label(rtop, text="④ 参照表:").grid(row=0, column=4, sticky=tk.W)
        self.rp_ref_var = tk.StringVar()
        self.rp_ref_box = ttk.Combobox(rtop, textvariable=self.rp_ref_var, state="readonly", width=34)
        self.rp_ref_box.grid(row=0, column=5, sticky=tk.W, padx=4)

        rp2 = ttk.Frame(self.nb2, padding=8)
        self.nb2.add(rp2, text="按字段匹配")
        ttk.Label(rp2, text="② 参照表格:").grid(row=0, column=0, sticky=tk.W)
        self.rp2_sheet_var = tk.StringVar()
        self.rp2_sheet_box = ttk.Combobox(rp2, textvariable=self.rp2_sheet_var, state="readonly", width=34)
        self.rp2_sheet_box.grid(row=0, column=1, sticky=tk.W, padx=4)
        self.rp2_sheet_box.bind("<<ComboboxSelected>>", lambda e: self._rp2_on_sheet())
        ttk.Label(rp2, text="③ 匹配字段:").grid(row=0, column=2, sticky=tk.W, padx=(12, 0))
        self.rp2_match_var = tk.StringVar()
        self.rp2_match_box = ttk.Combobox(rp2, textvariable=self.rp2_match_var, state="readonly", width=16)
        self.rp2_match_box.grid(row=0, column=3, sticky=tk.W, padx=4)
        ttk.Label(rp2, text="④ 取值字段:").grid(row=0, column=4, sticky=tk.W, padx=(12, 0))
        self.rp2_value_var = tk.StringVar()
        self.rp2_value_box = ttk.Combobox(rp2, textvariable=self.rp2_value_var, state="readonly", width=16)
        self.rp2_value_box.grid(row=0, column=5, sticky=tk.W, padx=4)
        ttk.Label(rp2, text="（目标子表需有与匹配字段同名的列，按它对齐行后把取值字段的值写入被替换列）",
                  foreground="#888888").grid(row=1, column=0, columnspan=6, sticky=tk.W, pady=(4, 0))

        # ---------- 共享操作栏和日志 ----------
        bar = ttk.Frame(self, padding=8)
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="登记字段", command=self._open_add_dialog).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(bar, text="删除映射组", command=self._open_delete_dialog).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(bar, text="更换工具表", command=self._change_tool_file).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(bar, text="设置备份目录", command=self._set_backup_dir).pack(side=tk.LEFT, padx=(0, 12))
        self.backup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="更新备份", variable=self.backup_var).pack(side=tk.LEFT)
        self.exec_btn = ttk.Button(bar, text="开始同步", command=lambda: self._run(dry=False))
        self.exec_btn.pack(side=tk.RIGHT, padx=4)
        ttk.Button(bar, text="预览", command=lambda: self._run(dry=True)).pack(side=tk.RIGHT)

        self.log_area = scrolledtext.ScrolledText(self, height=18, font=("Consolas", 9))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._on_tab_change())
        self._refresh_grp_values()
        self._refresh_rp_lists()

    def _refresh_grp_values(self):
        # 至少在 2 个表格里出现的映射组才可同步
        self.grp_box["values"] = [g for g in sorted(self.groups) if len(self.groups[g]) >= 2]

    def _field_label(self, t, g):
        """字段位置标签: 文件 | 表格 | 字段名"""
        return f"{t[0]} | {t[1]} | {self.groups[g][t]}"

    def _on_grp_change(self):
        g = self.grp_var.get()
        self._src_candidates = sorted(self.groups.get(g, {}))
        self.src_box["values"] = [self._field_label(t, g) for t in self._src_candidates]
        if self._src_candidates:
            self.src_box.current(0)
        self._on_src_change()

    def _on_src_change(self):
        g = self.grp_var.get()
        pos = self.src_box.current()
        self.tgt_list.delete(0, tk.END)
        self._tgt_candidates = []
        if pos < 0 or not g:
            self._refresh_key_candidates()
            return
        src_t = self._src_candidates[pos]
        self._tgt_candidates = [t for t in self._src_candidates if t != src_t]
        for t in self._tgt_candidates:
            self.tgt_list.insert(tk.END, self._field_label(t, g))
        self._refresh_key_candidates()

    def _select_all_targets(self):
        self.tgt_list.select_set(0, tk.END)
        self._refresh_key_candidates()

    def _involved_tables(self):
        """当前参与同步的表格: 同步源 + 已勾选的被同步字段(未勾选时按全部候选算)。"""
        pos = self.src_box.current()
        if pos < 0 or not getattr(self, "_src_candidates", None):
            return []
        src_t = self._src_candidates[pos]
        sel = self.tgt_list.curselection()
        if sel:
            return [src_t] + [self._tgt_candidates[i] for i in sel]
        return [src_t] + list(self._tgt_candidates)

    def _refresh_key_candidates(self):
        """键映射组候选: 能为所有参与表格提供对齐字段的组(排除同步组本身)。"""
        g = self.grp_var.get()
        involved = self._involved_tables()
        old = self.key_var.get()
        vals = []
        if g and involved:
            vals = [
                k for k in sorted(self.groups)
                if k != g and all(t in self.groups[k] for t in involved)
            ]
        self.key_box["values"] = vals
        if old in vals:
            self.key_var.set(old)
        else:
            self.key_var.set(vals[0] if vals else "")

    def _refresh_after_reload(self):
        """工具表内容变化后刷新界面，尽量保留当前选择。"""
        g, key = self.grp_var.get(), self.key_var.get()
        self._refresh_grp_values()
        if g in self.grp_box["values"]:
            self.grp_var.set(g)
            self._on_grp_change()
            if key in self.key_box["values"]:
                self.key_var.set(key)
        else:
            self.grp_var.set("")
            self.src_var.set("")
            self.src_box["values"] = []
            self.tgt_list.delete(0, tk.END)
            self.key_var.set("")
            self.key_box["values"] = []
        self._refresh_rp_lists()

    def _handle_lost_mappings(self, lost):
        """把失效映射完整列到日志区，并询问用户是否清理。"""
        if not lost:
            return
        self._log(f"[失效映射明细] 共 {len(lost)} 条（文件按工具表所在目录查找）：")
        for g, f, s, c, why in lost:
            self._log(f"  失效: {g} | {f} | {s} | {c}（{why}）")
        preview = "\n".join(f"  {g} | {f} | {s} | {c}（{why}）" for g, f, s, c, why in lost[:10])
        if len(lost) > 10:
            preview += f"\n  …… 共 {len(lost)} 条，完整清单见日志区"
        if messagebox.askyesno(
            "发现失效映射",
            f"有 {len(lost)} 条字段映射已失效：\n\n{preview}\n\n"
            "是否从字段表中删除这些映射？",
            parent=self):
            n = remove_mappings(self.tool_path, lost)
            self._reload()
            self._refresh_after_reload()
            self._log(f"[已清理] 删除 {n} 条失效映射")
        else:
            self._log("[提示] 失效映射未删除（同步时这些映射会报错或跳过）")

    def _change_tool_file(self):
        p = self._ensure_tool_file(force=True)
        if not p:
            self._log("[更换工具表] 已取消，保持当前工具表")
            return
        if p == self.tool_path:
            self._log(f"[更换工具表] 选择的仍是当前工具表「{p.name}」，未变化")
            return
        self.tool_path = p
        self.base_dir = p.resolve().parent
        # 备份目录: 用户显式设置过的保留，否则跟随新工具表
        self.backup_dir = find_backup_config(self.app_dir) or (self.base_dir / BACKUP_DIR)
        self.title(f"字段映射同步工具 - {p.name}")
        if self._reload():
            self._refresh_after_reload()
            self._log(f"[已切换工具表] {p.name}")
            self._startup_report()

    def _open_delete_dialog(self):
        def after_delete():
            if self._reload():
                self._refresh_after_reload()
        DeleteGroupDialog(self, after_delete)

    def _open_add_dialog(self):
        def after_save():
            if self._reload():
                self._refresh_after_reload()
                self._log("[工具表已更新] 登记数据已重新加载")
                self._startup_report()
        AddDialog(self, after_save)

    def _log(self, msg):
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)
        self.update_idletasks()

    def _on_tab_change(self):
        self.exec_btn.config(
            text="开始同步" if self.nb.index(self.nb.select()) == 0 else "开始替换")

    # ---------- 字段替换页签 ----------

    def _refresh_rp_lists(self):
        all_groups = sorted(self.groups)
        self.rp_src_box["values"] = all_groups
        self.rp_dst_box["values"] = all_groups
        # 字段池 = 所有表格的全部字段(含未登记), 保证替换功能通用
        self._rp_tables = scan_tables(self.base_dir, self.tool_path, self.backup_dir,
                                      cache_dir=self.app_dir)
        self._rp_all = [
            (f, s, col)
            for f, sheets in self._rp_tables
            for s, hdr in sheets
            for col in hdr
        ]
        self._rp2_flat = [(f, s, hdr) for f, sheets in self._rp_tables for s, hdr in sheets]
        self.rp2_sheet_box["values"] = [f"{f} | {s}" for f, s, _ in self._rp2_flat]
        # 清理已失效的已选字段
        if hasattr(self, "_rp_selected"):
            self._rp_selected &= set(self._rp_all)
        else:
            self._rp_selected = set()
        self._rp_files = sorted({it[0] for it in self._rp_all})
        self.rp_file_tree.delete(*self.rp_file_tree.get_children())
        build_file_tree(self.rp_file_tree, self._rp_files)
        self.rp_sheet_list.delete(0, tk.END)
        self._apply_rp_filter()
        self._update_rp_count()
        self._refresh_rp_ref()

    def _rp_label(self, item):
        f, s, col = item
        return f"{f} | {s} | {col}"

    def _rp_on_file(self):
        self.rp_sheet_list.delete(0, tk.END)
        f = tree_selected_file(self.rp_file_tree)
        if f:
            for s in sorted({it[1] for it in self._rp_all if it[0] == f}):
                self.rp_sheet_list.insert(tk.END, s)
        self._apply_rp_filter()

    def _rp_on_sheet(self):
        self._apply_rp_filter()

    def _rp_current_scope(self):
        """当前层级范围: 选中文件则限到该文件(再选子表则限到子表),
        选中文件夹则限到该文件夹子树, 都没选为全部。"""
        items = self._rp_all
        prefix = tree_selected_prefix(self.rp_file_tree)
        if prefix:
            if prefix.endswith("/"):
                # 文件夹节点: 整棵子树
                items = [it for it in items if it[0].startswith(prefix)]
            else:
                # 文件叶子
                items = [it for it in items if it[0] == prefix]
                si = self.rp_sheet_list.curselection()
                if si:
                    s = self.rp_sheet_list.get(si[0])
                    items = [it for it in items if it[1] == s]
        return items

    def _apply_rp_filter(self):
        """有关键字: 平铺全部匹配项(完整路径); 无关键字: 显示当前层级范围的字段。"""
        kw = self.rp_filter.get().strip().lower()
        if kw:
            self._rp_shown = [it for it in self._rp_all if kw in self._rp_label(it).lower()]
            labels = [self._rp_label(it) for it in self._rp_shown]
        else:
            self._rp_shown = self._rp_current_scope()
            # 选中具体文件时只显示字段名, 其余情况显示完整路径
            leaf = tree_selected_file(self.rp_file_tree) is not None
            labels = [it[2] if leaf else self._rp_label(it) for it in self._rp_shown]
        self.rp_list.delete(0, tk.END)
        for it, lbl in zip(self._rp_shown, labels):
            self.rp_list.insert(tk.END, lbl)
            if it in self._rp_selected:
                self.rp_list.select_set(tk.END)

    def _rp_on_select(self):
        """把当前列表的选中状态合并进全局已选集合（支持跨表累积选择）。"""
        shown = set(getattr(self, "_rp_shown", []))
        now = {self._rp_shown[i] for i in self.rp_list.curselection()}
        self._rp_selected -= shown
        self._rp_selected |= now
        self._update_rp_count()

    def _rp_select_all_shown(self):
        self.rp_list.select_set(0, tk.END)
        self._rp_on_select()

    def _rp_clear_selected(self):
        self._rp_selected.clear()
        self._apply_rp_filter()
        self._update_rp_count()

    def _update_rp_count(self):
        self.rp_count.config(text=f"已选 {len(self._rp_selected)} 个字段")

    def _rp2_on_sheet(self):
        pos = self.rp2_sheet_box.current()
        if pos < 0:
            return
        hdr = self._rp2_flat[pos][2]
        self.rp2_match_box["values"] = hdr
        self.rp2_value_box["values"] = hdr
        self.rp2_match_var.set("")
        self.rp2_value_var.set("")

    def _refresh_rp_ref(self):
        s, d = self.rp_src_var.get(), self.rp_dst_var.get()
        cands = []
        if s and d and s != d:
            cands = sorted(set(self.groups.get(s, {})) & set(self.groups.get(d, {})))
        self._rp_ref_candidates = cands
        self.rp_ref_box["values"] = [
            f"{t[0]} | {t[1]}（{self.groups[s][t]} → {self.groups[d][t]}）" for t in cands
        ]
        self.rp_ref_var.set("")
        if cands:
            self.rp_ref_box.current(0)

    def _run(self, dry):
        if self.nb.index(self.nb.select()) == 0:
            self._run_sync(dry)
        elif self.nb2.index(self.nb2.select()) == 0:
            self._run_replace(dry)
        else:
            self._run_replace_align(dry)

    def _run_sync(self, dry):
        g = self.grp_var.get()
        if not g:
            messagebox.showwarning("缺少选择", "请先选择同步映射组")
            return
        pos = self.src_box.current()
        if pos < 0:
            messagebox.showwarning("缺少选择", "请选择同步源字段")
            return
        src = self._src_candidates[pos]
        targets = [self._tgt_candidates[i] for i in self.tgt_list.curselection()]
        if not targets:
            messagebox.showwarning("缺少选择", "请至少勾选一个被同步字段")
            return
        key = self.key_var.get()
        if not key:
            messagebox.showwarning("缺少选择", "请选择键映射组（用于对齐行）")
            return
        if not dry and not messagebox.askyesno(
            "确认同步",
            f"同步映射组: {g}\n同步源: {self._field_label(src, g)}\n"
            f"被同步字段: {len(targets)} 个\n键映射组: {key}\n\n确定执行？"):
            return
        self._log("=" * 60)
        try:
            run_sync(self.base_dir, self.groups, key, src, [g], targets,
                     self.backup_dir if self.backup_var.get() else None, dry, self._log)
        except Exception as ex:
            self._log(f"[错误] {ex}")

    def _run_replace(self, dry):
        if not self._rp_selected:
            messagebox.showwarning("缺少选择", "请先勾选被替换字段")
            return
        # 同一列可能被选入多个组, 去重后执行
        targets = list(dict.fromkeys(
            it[:3] for it in sorted(self._rp_selected, key=self._rp_label)))
        s, d = self.rp_src_var.get(), self.rp_dst_var.get()
        if not s or not d:
            messagebox.showwarning("缺少选择", "请选择「当前值是」和「替换成」两个映射组")
            return
        if s == d:
            messagebox.showwarning("选择有误", "源映射组和目标映射组相同，无需替换")
            return
        pos = self.rp_ref_box.current()
        if pos < 0:
            messagebox.showwarning("缺少选择", "请选择参照表")
            return
        ref = self._rp_ref_candidates[pos]
        if not dry and not messagebox.askyesno(
            "确认替换",
            f"被替换字段: {len(targets)} 个\n翻译规则: 「{s}」→「{d}」\n"
            f"参照表: {ref[0]} | {ref[1]}\n\n只会改写数据区，表头不动。确定执行？"):
            return
        self._log("=" * 60)
        try:
            run_replace(self.base_dir, self.groups, s, d, ref, targets,
                        self.backup_dir if self.backup_var.get() else None, dry, self._log)
        except Exception as ex:
            self._log(f"[错误] {ex}")
    def _run_replace_align(self, dry):
        if not self._rp_selected:
            messagebox.showwarning("缺少选择", "请先勾选被替换字段")
            return
        targets = list(dict.fromkeys(
            it[:3] for it in sorted(self._rp_selected, key=self._rp_label)))
        pos = self.rp2_sheet_box.current()
        if pos < 0:
            messagebox.showwarning("缺少选择", "请选择参照表格")
            return
        f, s, hdr = self._rp2_flat[pos]
        m, v = self.rp2_match_var.get(), self.rp2_value_var.get()
        if not m or not v:
            messagebox.showwarning("缺少选择", "请选择匹配字段和取值字段")
            return
        if m == v:
            messagebox.showwarning("选择有误", "匹配字段和取值字段不能相同")
            return
        if not dry and not messagebox.askyesno(
            "确认替换",
            f"被替换字段: {len(targets)} 个\n参照表: {f} | {s}\n"
            f"按「{m}」对齐行，把「{v}」的值写入被替换列\n\n只会改写数据区，表头不动。确定执行？"):
            return
        self._log("=" * 60)
        try:
            run_replace_align(self.base_dir, (f, s), m, v, targets,
                              self.backup_dir if self.backup_var.get() else None, dry, self._log)
        except Exception as ex:
            self._log(f"[错误] {ex}")


if __name__ == "__main__":
    app = App()
    if app.winfo_exists():
        app.mainloop()
