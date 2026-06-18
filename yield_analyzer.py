#!/usr/bin/env python3
"""SLT Yield 分析工具：读取 slt_yield 格式 Excel，按工序合并并可视化月度良率。"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import json

import pandas as pd
import matplotlib

matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = ["PingFang SC", "Heiti SC", "SimHei", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False

REQUIRED_COLUMNS = ["工序", "PassQty", "TestQty", "FailQty", "YearMonth"]
TEMP_LABELS = ("常温", "低温", "高温")
CATEGORY_COLORS = {"常温": "#4472C4", "低温": "#70AD47", "高温": "#ED7D31"}
CONFIG_PATH = Path(__file__).parent / "process_categories.json"
DEFAULT_CATEGORY_MAP = {
    "MT1": "常温",
    "MT2": "常温",
    "MT3": "常温",
    "MT9": "低温",
    "MT10": "高温",
    "MT11": "高温",
}


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {"last_input": "", "files": {}}


def save_config(config: dict) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def file_config_key(filepath: str) -> str:
    return str(Path(filepath).resolve())


def get_saved_categories(filepath: str, config: dict) -> dict[str, str]:
    return config.get("files", {}).get(file_config_key(filepath), {})


def save_categories_for_file(filepath: str, mapping: dict[str, str], config: dict) -> None:
    key = file_config_key(filepath)
    config.setdefault("files", {})[key] = mapping
    config["last_input"] = key
    save_config(config)


def sort_processes(processes: list[str]) -> list[str]:
    def key(name: str):
        digits = "".join(c for c in name if c.isdigit())
        return (name[:2] if len(name) >= 2 else name, int(digits) if digits else 0, name)

    return sorted(processes, key=key)


def read_yield_file(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath, engine="openpyxl")
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"缺少必要列: {', '.join(missing)}")

    df = df.dropna(subset=["工序", "YearMonth"]).copy()
    for col in ("PassQty", "TestQty", "FailQty"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["PassQty", "TestQty"])
    df["PassQty"] = df["PassQty"].astype(int)
    df["TestQty"] = df["TestQty"].astype(int)
    if df["FailQty"].isna().all():
        df["FailQty"] = df["TestQty"] - df["PassQty"]
    else:
        df["FailQty"] = df["FailQty"].fillna(df["TestQty"] - df["PassQty"]).astype(int)

    df["YearMonth"] = df["YearMonth"].astype(str)
    return df


def merge_by_process(df: pd.DataFrame, process_category: dict[str, str]) -> pd.DataFrame:
    grouped = (
        df.groupby(["YearMonth", "工序"], as_index=False)
        .agg(PassQty=("PassQty", "sum"), TestQty=("TestQty", "sum"), FailQty=("FailQty", "sum"))
        .sort_values(["YearMonth", "工序"])
    )
    grouped["Yield"] = grouped["PassQty"] / grouped["TestQty"]
    grouped["FDPPM"] = grouped["FailQty"] / grouped["TestQty"] * 1_000_000
    grouped["温度类型"] = grouped["工序"].map(process_category)
    return grouped


def calc_category_yield(month_data: pd.DataFrame, category: str) -> float | None:
    cat_data = month_data[month_data["温度类型"] == category]
    if cat_data.empty or cat_data["TestQty"].sum() == 0:
        return None
    return cat_data["PassQty"].sum() / cat_data["TestQty"].sum()


def calc_total_yield(month_data: pd.DataFrame) -> float | None:
    """总Yield = 常温Yield × 低温Yield × 高温Yield（各温度下 Qty 合并后计算）。"""
    total = 1.0
    for cat in TEMP_LABELS:
        y = calc_category_yield(month_data, cat)
        if y is None:
            return None
        total *= y
    return total


def build_monthly_summary(merged: pd.DataFrame) -> pd.DataFrame:
    months = sorted(merged["YearMonth"].unique())
    processes = sort_processes(merged["工序"].unique().tolist())
    rows = []

    for ym in months:
        month_data = merged[merged["YearMonth"] == ym]
        row = {"YearMonth": ym}
        for proc in processes:
            proc_row = month_data[month_data["工序"] == proc]
            if proc_row.empty:
                row[proc] = None
            else:
                row[proc] = proc_row.iloc[0]["Yield"]
        row["总Yield"] = calc_total_yield(month_data)
        rows.append(row)

    return pd.DataFrame(rows)


def build_category_summary(merged: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for ym in sorted(merged["YearMonth"].unique()):
        month_data = merged[merged["YearMonth"] == ym]
        row = {"YearMonth": ym}
        for cat in TEMP_LABELS:
            cat_data = month_data[month_data["温度类型"] == cat]
            if cat_data.empty:
                row[f"{cat}PassQty"] = 0
                row[f"{cat}TestQty"] = 0
                row[f"{cat}Yield"] = None
            else:
                pass_sum = cat_data["PassQty"].sum()
                test_sum = cat_data["TestQty"].sum()
                row[f"{cat}PassQty"] = pass_sum
                row[f"{cat}TestQty"] = test_sum
                row[f"{cat}Yield"] = pass_sum / test_sum if test_sum else None
        rows.append(row)
    return pd.DataFrame(rows)


def build_temp_qty_merged(merged: pd.DataFrame) -> pd.DataFrame:
    """同温度工序按月份合并 PassQty / TestQty / FailQty。"""
    grouped = (
        merged.groupby(["YearMonth", "温度类型"], as_index=False)
        .agg(PassQty=("PassQty", "sum"), TestQty=("TestQty", "sum"), FailQty=("FailQty", "sum"))
    )
    grouped["Yield"] = grouped["PassQty"] / grouped["TestQty"]
    grouped["FDPPM"] = grouped["FailQty"] / grouped["TestQty"] * 1_000_000
    cat_order = {cat: i for i, cat in enumerate(TEMP_LABELS)}
    grouped["_sort"] = grouped["温度类型"].map(cat_order)
    grouped = grouped.sort_values(["YearMonth", "_sort"]).drop(columns="_sort")
    return grouped[["YearMonth", "温度类型", "PassQty", "TestQty", "FailQty", "Yield", "FDPPM"]]


def export_excel(
    merged: pd.DataFrame,
    monthly: pd.DataFrame,
    category: pd.DataFrame,
    temp_qty: pd.DataFrame,
    output_path: str,
):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="工序汇总", index=False)
        monthly.to_excel(writer, sheet_name="月度良率", index=False)
        category.to_excel(writer, sheet_name="温度分类汇总", index=False)
        temp_qty.to_excel(writer, sheet_name="温度Qty合并", index=False)

        from openpyxl.styles import Font

        bold = Font(bold=True)
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for cell in ws[1]:
                cell.font = bold
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    header = ws.cell(1, cell.column).value
                    if header and ("Yield" in str(header) or header == "总Yield"):
                        cell.number_format = "0.00%"
                    elif header and "FDPPM" in str(header):
                        cell.number_format = "#,##0"


class YieldAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SLT Yield 分析工具")
        self.geometry("1200x820")
        self.minsize(1000, 700)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.processes: list[str] = []
        self.process_vars: dict[str, tk.StringVar] = {}
        self.merged_df: pd.DataFrame | None = None
        self.monthly_df: pd.DataFrame | None = None
        self.config = load_config()
        self._loading = False

        self._build_ui()
        self._restore_last_session()

    def _build_ui(self):
        main = ttk.Frame(self, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main, text="文件选择", padding=8)
        file_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(file_frame, text="输入文件:").grid(row=0, column=0, sticky=tk.W, padx=(0, 4))
        ttk.Entry(file_frame, textvariable=self.input_path, width=70).grid(row=0, column=1, sticky=tk.EW, padx=4)
        ttk.Button(file_frame, text="浏览...", command=self._browse_input).grid(row=0, column=2, padx=4)

        ttk.Label(file_frame, text="输出文件:").grid(row=1, column=0, sticky=tk.W, padx=(0, 4), pady=(6, 0))
        ttk.Entry(file_frame, textvariable=self.output_path, width=70).grid(
            row=1, column=1, sticky=tk.EW, padx=4, pady=(6, 0)
        )
        ttk.Button(file_frame, text="浏览...", command=self._browse_output).grid(row=1, column=2, padx=4, pady=(6, 0))
        file_frame.columnconfigure(1, weight=1)

        self.process_frame = ttk.LabelFrame(main, text="工序分类（每道工序须且仅能归属一种温度类型）", padding=8)
        self.process_frame.pack(fill=tk.X, pady=(0, 8))

        self.process_inner = ttk.Frame(self.process_frame)
        self.process_inner.pack(fill=tk.X)
        ttk.Label(self.process_inner, text="请先选择输入文件以加载工序列表", foreground="gray").pack(anchor=tk.W)

        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(btn_frame, text="加载工序", command=self._load_processes).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="生成并导出", command=self._generate).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="刷新图表", command=self._refresh_chart).pack(side=tk.LEFT)

        chart_frame = ttk.LabelFrame(main, text="月度良率可视化", padding=4)
        chart_frame.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(chart_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.fig_all = Figure(figsize=(10, 4), dpi=100)
        self.canvas_all = FigureCanvasTkAgg(self.fig_all, master=self.notebook)
        self.notebook.add(self.canvas_all.get_tk_widget(), text="全部工序")

        self.fig_cat = Figure(figsize=(10, 4), dpi=100)
        self.canvas_cat = FigureCanvasTkAgg(self.fig_cat, master=self.notebook)
        self.notebook.add(self.canvas_cat.get_tk_widget(), text="温度分类")

        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, pady=(4, 0))

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="选择输入 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.input_path.set(path)
            default_out = Path(path).with_name(Path(path).stem + "_output.xlsx")
            self.output_path.set(str(default_out))
            self._load_processes()

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="选择输出 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def _restore_last_session(self):
        last = self.config.get("last_input", "")
        if last and Path(last).exists():
            self.input_path.set(last)
            self.output_path.set(str(Path(last).with_name(Path(last).stem + "_output.xlsx")))
            self._load_processes()

    def _on_category_changed(self, *_args):
        if self._loading:
            return
        path = self.input_path.get().strip()
        if not path or not self.process_vars:
            return
        mapping = {proc: self.process_vars[proc].get() for proc in self.processes}
        save_categories_for_file(path, mapping, self.config)

    def _load_processes(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择输入文件")
            return
        try:
            df = read_yield_file(path)
            self.processes = sort_processes(df["工序"].astype(str).unique().tolist())
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return

        for widget in self.process_inner.winfo_children():
            widget.destroy()

        if not self.processes:
            ttk.Label(self.process_inner, text="未找到有效工序数据", foreground="red").pack(anchor=tk.W)
            return

        header = ttk.Frame(self.process_inner)
        header.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(header, text="工序", width=12, font=("", 10, "bold")).pack(side=tk.LEFT, padx=4)
        for cat in TEMP_LABELS:
            ttk.Label(header, text=cat, width=10, font=("", 10, "bold")).pack(side=tk.LEFT, padx=8)

        self.process_vars.clear()
        saved = get_saved_categories(path, self.config)
        self._loading = True

        for proc in self.processes:
            row = ttk.Frame(self.process_inner)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=proc, width=12).pack(side=tk.LEFT, padx=4)
            if proc in saved and saved[proc] in TEMP_LABELS:
                initial = saved[proc]
            else:
                initial = DEFAULT_CATEGORY_MAP.get(proc, "常温")
            var = tk.StringVar(value=initial)
            var.trace_add("write", self._on_category_changed)
            self.process_vars[proc] = var
            for cat in TEMP_LABELS:
                ttk.Radiobutton(row, text="", variable=var, value=cat).pack(side=tk.LEFT, padx=28)

        self._loading = False
        save_categories_for_file(
            path,
            {proc: self.process_vars[proc].get() for proc in self.processes},
            self.config,
        )

        saved_hint = "（已恢复上次分类）" if saved else ""
        self.status_var.set(f"已加载 {len(self.processes)} 个工序: {', '.join(self.processes)}{saved_hint}")

    def _validate_categories(self) -> dict[str, str] | None:
        if not self.processes:
            messagebox.showwarning("提示", "请先加载工序")
            return None

        mapping = {proc: self.process_vars[proc].get() for proc in self.processes}
        for cat in TEMP_LABELS:
            if not any(v == cat for v in mapping.values()):
                messagebox.showerror("分类错误", f"「{cat}」工序至少需要选择一个")
                return None
        return mapping

    def _generate(self):
        input_path = self.input_path.get().strip()
        output_path = self.output_path.get().strip()
        if not input_path or not output_path:
            messagebox.showwarning("提示", "请选择输入和输出文件路径")
            return

        mapping = self._validate_categories()
        if mapping is None:
            return

        save_categories_for_file(input_path, mapping, self.config)

        try:
            df = read_yield_file(input_path)
            merged = merge_by_process(df, mapping)
            monthly = build_monthly_summary(merged)
            category = build_category_summary(merged)
            temp_qty = build_temp_qty_merged(merged)
            export_excel(merged, monthly, category, temp_qty, output_path)

            self.merged_df = merged
            self.monthly_df = monthly
            self._refresh_chart()
            self.status_var.set(f"已导出: {output_path}  ({len(merged)} 行工序汇总)")
            messagebox.showinfo("完成", f"Excel 已导出至:\n{output_path}")
        except Exception as e:
            messagebox.showerror("处理失败", str(e))

    def _refresh_chart(self):
        if self.merged_df is None or self.monthly_df is None:
            if self.input_path.get().strip() and self.process_vars:
                mapping = self._validate_categories()
                if mapping is None:
                    return
                try:
                    df = read_yield_file(self.input_path.get().strip())
                    self.merged_df = merge_by_process(df, mapping)
                    self.monthly_df = build_monthly_summary(self.merged_df)
                except Exception:
                    return
            else:
                return

        self._draw_process_chart()
        self._draw_category_chart()

    def _draw_process_chart(self):
        self.fig_all.clear()
        ax = self.fig_all.add_subplot(111)
        monthly = self.monthly_df
        merged = self.merged_df
        x = range(len(monthly))
        x_labels = monthly["YearMonth"].tolist()

        for proc in sort_processes(merged["工序"].unique().tolist()):
            cat = merged.loc[merged["工序"] == proc, "温度类型"].iloc[0]
            color = CATEGORY_COLORS.get(cat, "#666666")
            y_vals = monthly[proc].tolist()
            ax.plot(x, y_vals, marker="o", markersize=4, label=proc, color=color, linewidth=1.5)

        ax.plot(
            x,
            monthly["总Yield"].tolist(),
            marker="s",
            markersize=5,
            label="总Yield",
            color="#C00000",
            linewidth=2.5,
            linestyle="--",
        )

        ax.set_xticks(list(x))
        ax.set_xticklabels(x_labels, rotation=45, ha="right")
        ax.set_ylabel("Yield")
        ax.set_xlabel("YearMonth")
        ax.set_title("各工序月度良率及总Yield（总Yield = 常温 × 低温 × 高温）")
        ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0))
        ax.legend(loc="upper right", fontsize=8, ncol=2)
        ax.grid(True, alpha=0.3)
        self.fig_all.tight_layout()
        self.canvas_all.draw()

    def _draw_category_chart(self):
        self.fig_cat.clear()
        ax = self.fig_cat.add_subplot(111)
        merged = self.merged_df
        monthly = self.monthly_df
        months = sorted(merged["YearMonth"].unique())
        x = list(range(len(months)))

        for cat in TEMP_LABELS:
            yields = []
            for ym in months:
                cat_data = merged[(merged["YearMonth"] == ym) & (merged["温度类型"] == cat)]
                if cat_data.empty or cat_data["TestQty"].sum() == 0:
                    yields.append(None)
                else:
                    yields.append(cat_data["PassQty"].sum() / cat_data["TestQty"].sum())
            ax.plot(
                x,
                yields,
                marker="o",
                markersize=5,
                label=f"{cat}工序",
                color=CATEGORY_COLORS[cat],
                linewidth=2,
            )
            for xi, y in zip(x, yields):
                if y is not None:
                    ax.annotate(
                        f"{y * 100:.1f}%",
                        (xi, y),
                        textcoords="offset points",
                        xytext=(0, 8),
                        ha="center",
                        fontsize=7,
                        color=CATEGORY_COLORS[cat],
                    )

        total_yields = monthly["总Yield"].tolist()
        ax.plot(
            x,
            total_yields,
            marker="s",
            markersize=5,
            label="总Yield",
            color="#C00000",
            linewidth=2.5,
            linestyle="--",
        )
        for xi, y in zip(x, total_yields):
            if y is not None:
                ax.annotate(
                    f"{y * 100:.1f}%",
                    (xi, y),
                    textcoords="offset points",
                    xytext=(0, -12),
                    ha="center",
                    fontsize=7,
                    color="#C00000",
                )

        ax.set_xticks(x)
        ax.set_xticklabels(months, rotation=45, ha="right")
        ax.set_ylabel("Yield")
        ax.set_xlabel("YearMonth")
        ax.set_title("常温 / 低温 / 高温 工序月度合并良率及总Yield（总Yield = 常温 × 低温 × 高温）")
        ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0))
        ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1), borderaxespad=0, fontsize=8)
        ax.grid(True, alpha=0.3)
        ax.margins(y=0.12)
        self.fig_cat.subplots_adjust(right=0.78)
        self.fig_cat.tight_layout(rect=[0, 0, 0.78, 1])
        self.canvas_cat.draw()


def main():
    app = YieldAnalyzerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
