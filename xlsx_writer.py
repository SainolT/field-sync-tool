# -*- coding: utf-8 -*-
"""xlsx 外科手术写入器

不用 openpyxl 整体读写（那会重写样式表、丢弃图表/缓存值等，可能破坏下游转表），
而是直接在 xlsx(zip) 内部做最小修改:
  · 只改写目标 sheet 的 XML 文本中的目标 <c> 单元格
  · 其它 zip 条目字节级原样复制
  · 字符串值用 inlineStr 写入, 不动 sharedStrings.xml
  · 单元格样式属性(s)原样保留, 只换值
"""
import os
import re
import shutil
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape, unescape

_XML_UNESCAPE_MAP = {"&quot;": '"', "&apos;": "'", "&lt;": "<", "&gt;": ">", "&amp;": "&"}


def _unescape(text):
    return unescape(text, _XML_UNESCAPE_MAP)


def col_letters(idx):
    """1-based 列号转列字母: 1->A, 27->AA"""
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_paths(xlsx_path):
    """解析工作簿, 返回 {sheet名: zip内路径(如 xl/worksheets/sheet1.xml)}"""
    with zipfile.ZipFile(xlsx_path) as z:
        wb = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    name2rid = {}
    for m in re.finditer(r"<sheet\b[^>]*>", wb):
        tag = m.group(0)
        name = re.search(r'\bname="([^"]*)"', tag)
        rid = re.search(r'\br:id="([^"]*)"', tag)
        if name and rid:
            name2rid[_unescape(name.group(1))] = rid.group(1)
    rid2target = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        rid = re.search(r'\bId="([^"]*)"', tag)
        tgt = re.search(r'\bTarget="([^"]*)"', tag)
        if rid and tgt:
            rid2target[rid.group(1)] = tgt.group(1)
    out = {}
    for name, rid in name2rid.items():
        t = rid2target.get(rid)
        if t:
            out[name] = t.lstrip("/") if t.startswith("/") else "xl/" + t
    return out


def _cell_xml(ref, s_attr, value):
    """构造单元格 XML。保留样式属性 s_attr(如 ' s="3"', 可为空串)。"""
    if isinstance(value, bool):
        return f'<c r="{ref}"{s_attr} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s_attr}><v>{value}</v></c>'
    text = escape(str(value))
    space = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t{space}>{text}</t></is></c>'


_CELL_RE_CACHE = {}


def _find_cell(row_xml, ref):
    """在行 XML 里找 <c r="ref"> 元素, 返回 (start, end, tag) 或 None。"""
    pat = _CELL_RE_CACHE.get(ref)
    if pat is None:
        pat = re.compile(r'<c\b(?=[^>]*\br="' + re.escape(ref) + r'")[^>]*?(?:/>|>.*?</c>)',
                         re.DOTALL)
        _CELL_RE_CACHE[ref] = pat
    m = pat.search(row_xml)
    if not m:
        return None
    tag_open = re.match(r"<c\b[^>]*", m.group(0)).group(0)
    return m.start(), m.end(), tag_open


def _edit_row(xml, row_num, changes):
    """修改指定行的若干单元格。changes: [(col_idx, new_value)]。
    返回 (新xml, applied数, missed列字母列表)。"""
    row_pat = re.compile(r'<row\b(?=[^>]*\br="' + str(row_num) + r'")[^>]*?(/>|>)')
    m = row_pat.search(xml)
    if not m:
        # 行不存在: 按行号顺序插入 <sheetData>
        new_cells = "".join(
            _cell_xml(f"{col_letters(c)}{row_num}", "", v) for c, v in sorted(changes))
        new_row = f'<row r="{row_num}">{new_cells}</row>'
        insert_at = None
        for rm in re.finditer(r'<row\b(?=[^>]*\br="(\d+)")', xml):
            if int(rm.group(1)) > row_num:
                insert_at = rm.start()
                break
        if insert_at is None:
            end = xml.find("</sheetData>")
            if end < 0:
                return xml, 0, [col_letters(c) for c, _ in changes]
            xml = xml[:end] + new_row + xml[end:]
        else:
            xml = xml[:insert_at] + new_row + xml[insert_at:]
        return xml, len(changes), []

    if m.group(1) == "/>":  # 空行元素 <row r="N"/>
        start, end = m.start(), m.end()
        row_xml = ""
        row_close = None
    else:
        start = m.start()
        close = xml.find("</row>", m.end())
        if close < 0:
            return xml, 0, [col_letters(c) for c, _ in changes]
        row_xml = xml[m.end():close]
        row_close = close

    applied, missed = 0, []
    for col_idx, value in sorted(changes):
        ref = f"{col_letters(col_idx)}{row_num}"
        found = _find_cell(row_xml, ref)
        if found:
            cs, ce, tag_open = found
            sm = re.search(r'\bs="(\d+)"', tag_open)
            s_attr = f' s="{sm.group(1)}"' if sm else ""
            row_xml = row_xml[:cs] + _cell_xml(ref, s_attr, value) + row_xml[ce:]
            applied += 1
        else:
            # 按列顺序插入新单元格
            insert_at = len(row_xml)
            for cm in re.finditer(r'<c\b(?=[^>]*\br="([A-Z]+)\d+")', row_xml):
                letters = cm.group(1)
                ci = 0
                for ch in letters:
                    ci = ci * 26 + (ord(ch) - 64)
                if ci > col_idx:
                    insert_at = cm.start()
                    break
            row_xml = row_xml[:insert_at] + _cell_xml(ref, "", value) + row_xml[insert_at:]
            applied += 1

    if m.group(1) == "/>":
        open_tag = m.group(0)[:-2] + ">"
        xml = xml[:start] + open_tag + row_xml + "</row>" + xml[end:]
    else:
        xml = xml[:m.end()] + row_xml + xml[row_close:]
    return xml, applied, missed


def apply_file_changes(xlsx_path, changes_by_sheet):
    """对一个 xlsx 做外科手术写入。
    changes_by_sheet: {sheet名: [(row, col_idx, new_value), ...]}
    返回 (应用的单元格数, [未命中的描述...])。调用方负责备份。"""
    xlsx_path = Path(xlsx_path)
    spaths = sheet_paths(xlsx_path)
    edits = {}
    missed = []
    applied = 0
    for sheet, changes in changes_by_sheet.items():
        spath = spaths.get(sheet)
        if not spath:
            missed.append(f"工作表「{sheet}」不存在")
            continue
        edits.setdefault(spath, []).extend(changes)
    if missed:
        return 0, missed

    new_contents = {}
    with zipfile.ZipFile(xlsx_path) as z:
        for spath, changes in edits.items():
            xml = z.read(spath).decode("utf-8")
            by_row = {}
            for r, c, v in changes:
                by_row.setdefault(r, []).append((c, v))
            for r in sorted(by_row):
                xml, ok, miss = _edit_row(xml, r, by_row[r])
                applied += ok
                missed.extend(f"{spath} 行{r} 列{mc}" for mc in miss)
            new_contents[spath] = xml
        # 写 zip: 除被改的 sheet XML 外, 其它条目字节级原样复制
        tmp = xlsx_path.with_name(xlsx_path.name + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in z.infolist():
                data = z.read(item.filename)
                if item.filename in new_contents:
                    data = new_contents[item.filename].encode("utf-8")
                zout.writestr(item, data)
    os.replace(tmp, xlsx_path)
    return applied, missed


def save_file_with_changes(base_dir, rel, changes_by_sheet, backup_dir, stamp):
    """备份 + 外科手术写入。返回备份文件名(未备份返回 None)。"""
    tp = Path(base_dir) / rel
    bak_name = None
    if backup_dir:
        bdir = Path(backup_dir)
        bdir.mkdir(parents=True, exist_ok=True)
        safe = "__".join(Path(rel).with_suffix("").parts)
        bak_name = f"{safe}.{stamp}{tp.suffix}"
        shutil.copy2(tp, bdir / bak_name)  # 写盘前备份, 磁盘上仍是原始文件
    applied, missed = apply_file_changes(tp, changes_by_sheet)
    return bak_name, applied, missed
