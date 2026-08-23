# -*- coding: utf-8 -*-
"""
快递单号比对审查工具 (macOS / Windows)
- 选择两份 Excel, 按快递单号匹配
- 单边存在 = 异常; 具体重量 vs 区间重量模糊比对
- 输出审查报告(明细/异常提取/双方标注)
"""
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from compare_engine import compare, export_report


def guess_columns(path):
    """读取表头, 猜测单号列和重量列"""
    try:
        df = pd.read_excel(path, sheet_name=0, nrows=0)
        cols = [str(c).strip() for c in df.columns]
    except Exception:
        return "", ""
    key = next((c for c in cols if "单号" in c or "运单" in c or "编号" in c), cols[0] if cols else "")
    weight = next((c for c in cols if "区间" in c) ,
                  next((c for c in cols if "重" in c), ""))
    return key, weight


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("快递单号比对审查工具")
        self.geometry("760x520")
        self.file_a = tk.StringVar()
        self.file_b = tk.StringVar()
        self._build()

    # ---------- 界面 ----------
    def _build(self):
        pad = {"padx": 10, "pady": 6}
        f1 = ttk.LabelFrame(self, text=" A 方表格（具体重量一方）")
        f1.pack(fill="x", **pad)
        ttk.Entry(f1, textvariable=self.file_a).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(f1, text="选择文件…", command=lambda: self.pick(self.file_a, suffix="_A")).pack(side="left", padx=8)

        f2 = ttk.LabelFrame(self, text=" B 方表格（区间重量一方）")
        f2.pack(fill="x", **pad)
        ttk.Entry(f2, textvariable=self.file_b).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(f2, text="选择文件…", command=lambda: self.pick(self.file_b, suffix="_B")).pack(side="left", padx=8)

        f3 = ttk.LabelFrame(self, text=" 列选择（选完文件后从下拉框中选择对应列）")
        f3.pack(fill="x", **pad)
        self._cols_a = []   # A 方实际列名列表
        self._cols_b = []
        self.key_a = tk.StringVar(value="请先选文件")
        self.weight_a = tk.StringVar(value="请先选文件")
        self.key_b = tk.StringVar(value="请先选文件")
        self.weight_b = tk.StringVar(value="请先选文件")
        grid = ttk.Frame(f3); grid.pack(padx=8, pady=8)
        specs = [("A方·单号列", self.key_a), ("A方·重量列", self.weight_a),
                 ("B方·单号列", self.key_b), ("B方·区间列", self.weight_b)]
        self.cboxes = {}
        for i, (t, v) in enumerate(specs):
            r, c = divmod(i, 2)
            ttk.Label(grid, text=t).grid(row=r, column=c * 2, sticky="e", padx=(0, 4), pady=3)
            cb = ttk.Combobox(grid, textvariable=v, width=18, state="readonly")
            cb.grid(row=r, column=c * 2 + 1, pady=3)
            self.cboxes[v] = cb

        f4 = ttk.LabelFrame(self, text=" 高级选项")
        f4.pack(fill="x", **pad)
        row = ttk.Frame(f4); row.pack(padx=8, pady=6)
        ttk.Label(row, text="重量容差(kg):").pack(side="left")
        self.tol = tk.StringVar(value="0")
        ttk.Entry(row, textvariable=self.tol, width=8).pack(side="left", padx=6)
        ttk.Label(row, text="(例如填 0.05 表示 ±50g 以内视为一致)").pack(side="left")

        ttk.Button(self, text="开始比对审查", command=self.run).pack(fill="x", padx=10, pady=8,
                                                                    ipady=6)

        f5 = ttk.LabelFrame(self, text=" 结果日志")
        f5.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(f5, height=10, state="disabled", font=("Menlo", 11))
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

    def pick(self, var, suffix=""):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls"), ("所有文件", "*.*")])
        if not path:
            return
        var.set(path)

        # 读取该文件所有列名, 填充下拉框并自动预选猜测项
        try:
            cols = [str(c).strip() for c in pd.read_excel(path, sheet_name=0, nrows=0).columns]
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败:\n{e}")
            return
        if suffix == "_A":
            self._cols_a = cols
            for var_ in (self.key_a, self.weight_a):
                self.cboxes[var_]["values"] = cols
            key, weight = guess_columns(path)
            if key:   self.key_a.set(key)
            if weight: self.weight_a.set(weight)
        else:
            self._cols_b = cols
            for var_ in (self.key_b, self.weight_b):
                self.cboxes[var_]["values"] = cols
            key, weight = guess_columns(path)
            if key:   self.key_b.set(key)
            if weight:
                # B方优先把含“区间”的列预选到区间框
                wcol = next((c for c in cols if "区间" in c), weight)
                self.weight_b.set(wcol)
        self.write_log(f"已加载{'A' if suffix == '_A' else 'B'}方: {os.path.basename(path)} "
                       f"(共 {len(cols)} 列, 请确认下方列选择)")

    # ---------- 逻辑 ----------
    def write_log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def run(self):
        fa, fb = self.file_a.get().strip(), self.file_b.get().strip()
        if not fa or not fb:
            messagebox.showwarning("提示", "请先选择两份 Excel 文件")
            return
        ka, wa, kb, wb = (self.key_a.get().strip(), self.weight_a.get().strip(),
                          self.key_b.get().strip(), self.weight_b.get().strip())
        missing = [name for name, v in [("A方·单号列", ka), ("A方·重量列", wa),
                                        ("B方·单号列", kb), ("B方·区间列", wb)]
                   if not v or v == "请先选文件"]
        if missing:
            messagebox.showwarning("提示", "请为以下项选择列:\n" + "\n".join(missing))
            return
        try:
            tol = float(self.tol.get() or 0)
        except ValueError:
            messagebox.showwarning("提示", "容差必须是数字")
            return
        self.write_log("=" * 50)
        self.write_log(f"A方: {os.path.basename(fa)}")
        self.write_log(f"B方: {os.path.basename(fb)}")
        self.write_log("正在比对…")
        threading.Thread(target=self._work, args=(fa, fb, tol), daemon=True).start()

    def _work(self, fa, fb, tol):
        try:
            results, *_ = compare(fa, fb,
                                  key_a=self.key_a.get().strip(),
                                  weight_a=self.weight_a.get().strip(),
                                  key_b=self.key_b.get().strip(),
                                  weight_b=self.weight_b.get().strip(),
                                  tolerance=tol)
            out = os.path.join(os.path.dirname(fa), "审查报告.xlsx")
            summary, out = export_report(results, fa, fb,
                                         self.key_a.get().strip(), self.key_b.get().strip(),
                                         self.weight_a.get().strip(), self.weight_b.get().strip(),
                                         out_path=out)
            total = len(results)
            abnormal = total - summary["比对一致"]
            self.after(0, self._done, summary, out, abnormal)
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            self.after(0, lambda: (self.write_log("❌ 出错: " + err),
                                   messagebox.showerror("错误", err)))

    def _done(self, summary, out, abnormal):
        self.write_log(f"✅ 完成! 一致 {summary['比对一致']} 条 | "
                       f"重量不符 {summary['重量不符']} | 仅A方有 {summary['仅A方有(B缺单)']} | "
                       f"仅B方有 {summary['仅B方有(A缺单)']}")
        self.write_log(f"报告已保存: {out}")
        if abnormal:
            messagebox.showwarning("审查完成", f"发现 {abnormal} 条异常!\n\n报告:\n{out}")
        else:
            messagebox.showinfo("审查完成", f"全部一致, 共 {summary['比对一致']} 条。\n\n报告:\n{out}")


if __name__ == "__main__":
    App().mainloop()
