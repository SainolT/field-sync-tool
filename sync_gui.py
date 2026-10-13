# -*- coding: utf-8 -*-
"""字段映射同步工具（窗口版）

界面层。全部业务逻辑在 core.py（工具表/扫描/匹配计划/执行编排），
Excel 写入在 xlsx_writer.py（外科手术式 XML 编辑，不动格式）。

运行: python sync_gui.py
"""
import tkinter as tk
from pathlib import Path
from tkinter import ttk, scrolledtext, messagebox, filedialog

from core import *  # 逻辑层: 常量/配置/工具表/扫描/计划与执行
from core import app_dir  # 显式引用, 防止 star 导入遗漏

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
        self.title(f"字段映射同步工具 v{VERSION}")
        self.geometry("1020x680")
        self.app_dir = app_dir()          # 配置文件所在目录（跟随 exe/脚本）
        self.base_dir = self.app_dir      # 数据目录（跟随工具表，选定后更新）
        self.tool_path = self._ensure_tool_file()
        if not self.tool_path:
            self.destroy()
            return
        self.base_dir = self.tool_path.resolve().parent
        self.backup_dir = find_backup_config(self.app_dir) or (self.base_dir / BACKUP_DIR)
        self.title(f"字段映射同步工具 v{VERSION} - {self.tool_path.name}")
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
        self.title(f"字段映射同步工具 v{VERSION} - {p.name}")
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
