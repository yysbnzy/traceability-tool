# -*- coding: utf-8 -*-
"""
测试用例与需求溯源工具 v3.1 - 优化UI布局
"""

import re
import sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class TraceabilityTool:
    def __init__(self, root):
        self.root = root
        self.root.title("测试用例与需求溯源工具 v3.1")
        self.root.geometry("900x650")
        self.root.minsize(850, 600)
        
        self.case_file_path = None
        self.req_file_path = None
        self.case_wb = None
        self.req_wb = None
        
        # 配置样式
        self.setup_styles()
        self.create_ui()
    
    def setup_styles(self):
        """配置ttk样式"""
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Microsoft YaHei', 11, 'bold'), foreground='#2c3e50')
        style.configure('Group.TLabelframe', font=('Microsoft YaHei', 10, 'bold'))
        style.configure('Group.TLabelframe.Label', font=('Microsoft YaHei', 10, 'bold'), foreground='#34495e')
        style.configure('Hint.TLabel', font=('Microsoft YaHei', 9), foreground='#7f8c8d')
    
    def create_ui(self):
        # 主容器 - 使用Canvas支持滚动
        main_container = ttk.Frame(self.root, padding="15")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # ===== 标题 =====
        title_frame = ttk.Frame(main_container)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(title_frame, text="📊 测试用例与需求溯源工具", style='Title.TLabel').pack(side=tk.LEFT)
        ttk.Label(title_frame, text="v3.1", style='Hint.TLabel').pack(side=tk.LEFT, padx=(10, 0))
        
        # ===== 用例文档配置（顶部通栏）=====
        case_doc_frame = ttk.LabelFrame(main_container, text="📄 用例文档配置", style='Group.TLabelframe', padding="10")
        case_doc_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 文件选择行
        file_row = ttk.Frame(case_doc_frame)
        file_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(file_row, text="文件路径:", width=10).pack(side=tk.LEFT)
        self.case_file_var = tk.StringVar()
        ttk.Entry(file_row, textvariable=self.case_file_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(file_row, text="📂 浏览", command=self.select_case_file, width=8).pack(side=tk.LEFT)
        
        # 列选择行
        col_row = ttk.Frame(case_doc_frame)
        col_row.pack(fill=tk.X)
        ttk.Label(col_row, text="用例ID列:", width=10).pack(side=tk.LEFT)
        self.case_id_col = ttk.Combobox(col_row, values=list("ABCDEFGHIJKLMNOPQRSTUVWXYZ"), width=8, state="readonly")
        self.case_id_col.pack(side=tk.LEFT, padx=(0, 20))
        self.case_id_col.set("E")
        
        ttk.Label(col_row, text="需求ID列:", width=10).pack(side=tk.LEFT)
        self.case_req_col = ttk.Combobox(col_row, values=list("ABCDEFGHIJKLMNOPQRSTUVWXYZ"), width=8, state="readonly")
        self.case_req_col.pack(side=tk.LEFT)
        self.case_req_col.set("B")
        
        # 存储控件引用（必须在create_concat_config之前初始化）
        self.concat_widgets = {}
        
        # ===== 中间左右分栏 =====
        middle_frame = ttk.Frame(main_container)
        middle_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        middle_frame.columnconfigure(0, weight=1)
        middle_frame.columnconfigure(1, weight=1)
        
        # 左侧：用例ID拼接
        left_frame = ttk.LabelFrame(middle_frame, text="🔗 用例ID拼接配置", style='Group.TLabelframe', padding="10")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        self.create_concat_config(left_frame, "case")
        
        # 右侧：需求ID拼接
        right_frame = ttk.LabelFrame(middle_frame, text="🔗 需求ID拼接配置", style='Group.TLabelframe', padding="10")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        
        self.create_concat_config(right_frame, "req")
        
        # ===== 需求文档配置（底部）=====
        req_doc_frame = ttk.LabelFrame(main_container, text="📄 需求文档配置（可选）", style='Group.TLabelframe', padding="10")
        req_doc_frame.pack(fill=tk.X, pady=(0, 15))
        
        req_row = ttk.Frame(req_doc_frame)
        req_row.pack(fill=tk.X)
        ttk.Label(req_row, text="文件路径:", width=10).pack(side=tk.LEFT)
        self.req_file_var = tk.StringVar()
        ttk.Entry(req_row, textvariable=self.req_file_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(req_row, text="📂 浏览", command=self.select_req_file, width=8).pack(side=tk.LEFT)
        
        ttk.Label(req_row, text="需求ID列:", width=10).pack(side=tk.LEFT, padx=(20, 0))
        self.req_id_col = ttk.Combobox(req_row, values=list("ABCDEFGHIJKLMNOPQRSTUVWXYZ"), width=8, state="readonly")
        self.req_id_col.pack(side=tk.LEFT)
        self.req_id_col.set("A")
        
        # ===== 底部按钮 =====
        btn_frame = ttk.Frame(main_container)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="🚀 开始溯源分析", command=self.start_analysis, width=20).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="❌ 退出", command=self.root.quit, width=10).pack(side=tk.RIGHT)
    
    def create_concat_config(self, parent, prefix):
        """创建拼接配置区域"""
        # 前缀配置
        prefix_frame = ttk.Frame(parent)
        prefix_frame.pack(fill=tk.X, pady=(0, 12))
        
        ttk.Label(prefix_frame, text="前缀来源:", width=10).pack(side=tk.LEFT)
        source_combo = ttk.Combobox(prefix_frame, values=["固定文本", "用例文档列", "需求文档列"], 
                                     width=12, state="readonly")
        source_combo.pack(side=tk.LEFT, padx=5)
        source_combo.set("固定文本")
        
        value_frame = ttk.Frame(prefix_frame)
        value_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        value_entry = ttk.Entry(value_frame, width=15)
        value_entry.pack(fill=tk.X, expand=True)
        
        self.concat_widgets[f"{prefix}_prefix_source"] = source_combo
        self.concat_widgets[f"{prefix}_prefix_value"] = value_entry
        self.concat_widgets[f"{prefix}_prefix_frame"] = value_frame
        
        source_combo.bind("<<ComboboxSelected>>", 
                         lambda e, p=prefix, t="prefix": self.on_source_change(p, t))
        
        # 后缀配置
        suffix_frame = ttk.Frame(parent)
        suffix_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(suffix_frame, text="后缀来源:", width=10).pack(side=tk.LEFT)
        source_combo2 = ttk.Combobox(suffix_frame, values=["固定文本", "用例文档列", "需求文档列"], 
                                      width=12, state="readonly")
        source_combo2.pack(side=tk.LEFT, padx=5)
        source_combo2.set("固定文本")
        
        value_frame2 = ttk.Frame(suffix_frame)
        value_frame2.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        value_entry2 = ttk.Entry(value_frame2, width=15)
        value_entry2.pack(fill=tk.X, expand=True)
        
        self.concat_widgets[f"{prefix}_suffix_source"] = source_combo2
        self.concat_widgets[f"{prefix}_suffix_value"] = value_entry2
        self.concat_widgets[f"{prefix}_suffix_frame"] = value_frame2
        
        source_combo2.bind("<<ComboboxSelected>>", 
                          lambda e, p=prefix, t="suffix": self.on_source_change(p, t))
        
        # 预览说明
        hint_frame = ttk.Frame(parent)
        hint_frame.pack(fill=tk.X, pady=(15, 0))
        ttk.Label(hint_frame, text="💡 示例: 前缀 + 原始ID + 后缀", 
                 style='Hint.TLabel', wraplength=350).pack()
    
    def on_source_change(self, prefix, field_type):
        """当来源类型改变时切换输入控件"""
        key = f"{prefix}_{field_type}"
        source_combo = self.concat_widgets[f"{key}_source"]
        value_frame = self.concat_widgets[f"{key}_frame"]
        
        source_type = source_combo.get()
        
        # 清除旧控件
        for widget in value_frame.winfo_children():
            widget.destroy()
        
        if source_type == "固定文本":
            new_entry = ttk.Entry(value_frame, width=15)
            new_entry.pack(fill=tk.X, expand=True)
            self.concat_widgets[f"{key}_value"] = new_entry
        else:
            col_combo = ttk.Combobox(value_frame, values=list("ABCDEFGHIJKLMNOPQRSTUVWXYZ"), 
                                     width=8, state="readonly")
            col_combo.pack(fill=tk.X, expand=True)
            col_combo.set("A")
            self.concat_widgets[f"{key}_value"] = col_combo
    
    def get_concat_value(self, prefix, field_type, case_row, req_row):
        """获取拼接值"""
        key = f"{prefix}_{field_type}"
        source = self.concat_widgets[f"{key}_source"].get()
        widget = self.concat_widgets[f"{key}_value"]
        
        if source == "固定文本":
            return widget.get().strip()
        elif source == "用例文档列":
            col_idx = self.col_to_index(widget.get())
            if case_row and len(case_row) > col_idx:
                val = case_row[col_idx]
                return str(val).strip() if val else ""
            return ""
        elif source == "需求文档列":
            if not self.req_wb:
                return None
            col_idx = self.col_to_index(widget.get())
            if req_row and len(req_row) > col_idx:
                val = req_row[col_idx]
                return str(val).strip() if val else ""
            return ""
        return ""
    
    def select_case_file(self):
        file_path = filedialog.askopenfilename(
            title="选择用例文档", 
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.case_file_path = file_path
            self.case_file_var.set(file_path)
            try:
                self.case_wb = load_workbook(file_path, read_only=True, data_only=True)
            except Exception as e:
                messagebox.showerror("错误", f"加载用例文档失败: {str(e)}")
    
    def select_req_file(self):
        file_path = filedialog.askopenfilename(
            title="选择需求文档", 
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.req_file_path = file_path
            self.req_file_var.set(file_path)
            try:
                self.req_wb = load_workbook(file_path, read_only=True, data_only=True)
            except Exception as e:
                messagebox.showerror("错误", f"加载需求文档失败: {str(e)}")
    
    def split_requirement_ids(self, text):
        if text is None:
            return []
        normalized = str(text)
        normalized = re.sub(r'[，；。\s\n]+', ',', normalized)
        parts = re.split(r'[,;]+', normalized)
        ids = [p.strip() for p in parts if p.strip()]
        return ids
    
    def col_to_index(self, col):
        return ord(col) - ord('A')
    
    def start_analysis(self):
        if not self.case_file_path:
            messagebox.showerror("错误", "请选择用例文档")
            return
        
        # 检查是否需要需求文档
        sources = [
            self.concat_widgets["case_prefix_source"].get(),
            self.concat_widgets["case_suffix_source"].get(),
            self.concat_widgets["req_prefix_source"].get(),
            self.concat_widgets["req_suffix_source"].get(),
        ]
        if "需求文档列" in sources and not self.req_wb:
            messagebox.showwarning("提示", "您选择了'需求文档列'作为来源，但未上传需求文档。\n请先上传需求文档，或修改来源为其他选项。")
            return
        
        case_id_idx = self.col_to_index(self.case_id_col.get())
        case_req_idx = self.col_to_index(self.case_req_col.get())
        req_id_idx = self.col_to_index(self.req_id_col.get())
        
        try:
            self.analyze(case_id_idx, case_req_idx, req_id_idx)
            messagebox.showinfo("完成", "✅ 溯源分析已完成！")
        except Exception as e:
            messagebox.showerror("错误", f"分析失败: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def analyze(self, case_id_idx, case_req_idx, req_id_idx):
        case_to_reqs = {}
        case_to_reqs_raw = {}
        req_to_cases = {}
        
        case_ws = self.case_wb.active
        max_col = max(case_id_idx, case_req_idx) + 1
        
        # 预加载需求文档数据
        req_data_cache = {}
        if self.req_wb:
            req_ws = self.req_wb.active
            for idx, row in enumerate(req_ws.iter_rows(min_row=2, values_only=True), start=2):
                req_data_cache[idx] = row
        
        for idx, row in enumerate(case_ws.iter_rows(min_row=2, max_col=max_col, values_only=True), start=2):
            if len(row) <= max(case_id_idx, case_req_idx):
                continue
            
            case_id_raw = str(row[case_id_idx]).strip() if row[case_id_idx] else ''
            req_text = row[case_req_idx]
            req_ids = self.split_requirement_ids(req_text)
            
            if (not case_id_raw or case_id_raw == 'None') and not req_ids:
                continue
            
            req_row_data = req_data_cache.get(idx)
            
            # 动态获取前缀后缀
            case_prefix = self.get_concat_value("case", "prefix", row, req_row_data)
            case_suffix = self.get_concat_value("case", "suffix", row, req_row_data)
            
            case_id_full = None
            if case_id_raw and case_id_raw != 'None':
                case_id_full = f"{case_prefix}{case_id_raw}{case_suffix}"
                if case_id_full not in case_to_reqs:
                    case_to_reqs[case_id_full] = set()
                    case_to_reqs_raw[case_id_raw] = set()
            
            for req_id in req_ids:
                req_prefix = self.get_concat_value("req", "prefix", row, req_row_data)
                req_suffix = self.get_concat_value("req", "suffix", row, req_row_data)
                req_id_full = f"{req_prefix}{req_id}{req_suffix}"
                
                if req_id_full not in req_to_cases:
                    req_to_cases[req_id_full] = set()
                if case_id_full:
                    req_to_cases[req_id_full].add(case_id_full)
                if case_id_full:
                    case_to_reqs[case_id_full].add(req_id_full)
                    case_to_reqs_raw[case_id_raw].add(req_id_full)
        
        # 需求文档所有需求
        all_req_ids_raw = set()
        if self.req_wb:
            req_ws = self.req_wb.active
            for row in req_ws.iter_rows(min_row=2, max_col=req_id_idx+1, values_only=True):
                if len(row) <= req_id_idx:
                    continue
                req_id = str(row[req_id_idx]).strip() if row[req_id_idx] else ''
                if req_id and req_id != 'None':
                    all_req_ids_raw.add(req_id)
        
        missing_case_reqs = all_req_ids_raw - set(req_to_cases.keys())
        
        output_path = filedialog.asksaveasfilename(
            title="保存分析结果", 
            defaultextension=".xlsx", 
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if not output_path:
            return
        
        self.export(output_path, case_to_reqs, case_to_reqs_raw, req_to_cases, 
                   all_req_ids_raw, missing_case_reqs)
    
    def export(self, output_path, case_to_reqs, case_to_reqs_raw, req_to_cases, 
               all_req_ids_raw, missing_case_reqs):
        wb = Workbook()
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                            top=Side(style='thin'), bottom=Side(style='thin'))
        
        # 双向溯源表
        ws1 = wb.active
        ws1.title = "双向溯源表"
        ws1['A1'] = '用例ID'
        ws1['B1'] = '关联需求ID'
        for cell in [ws1['A1'], ws1['B1']]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
        
        sorted_cases = sorted([c for c, r in case_to_reqs.items() if r])
        for idx, case_id in enumerate(sorted_cases, start=2):
            reqs = sorted(case_to_reqs[case_id])
            ws1.cell(row=idx, column=1, value=case_id)
            ws1.cell(row=idx, column=2, value=", ".join(reqs))
            ws1.cell(row=idx, column=1).border = thin_border
            ws1.cell(row=idx, column=2).border = thin_border
        
        ws1['D1'] = '需求ID'
        ws1['E1'] = '关联用例ID'
        for cell in [ws1['D1'], ws1['E1']]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
        
        sorted_reqs = sorted([r for r, c in req_to_cases.items() if c])
        for idx, req_id in enumerate(sorted_reqs, start=2):
            cases = sorted(req_to_cases[req_id])
            ws1.cell(row=idx, column=4, value=req_id)
            ws1.cell(row=idx, column=5, value=", ".join(cases))
            ws1.cell(row=idx, column=4).border = thin_border
            ws1.cell(row=idx, column=5).border = thin_border
        
        for col in ['A', 'B', 'D', 'E']:
            ws1.column_dimensions[col].width = 40
        
        # 输入源
        ws2 = wb.create_sheet("输入源")
        ws2['A1'] = '用例ID'
        ws2['B1'] = '需求ID'
        for cell in [ws2['A1'], ws2['B1']]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
        
        case_ws = self.case_wb.active
        max_col = max(self.col_to_index(self.case_id_col.get()), 
                     self.col_to_index(self.case_req_col.get())) + 1
        for idx, row in enumerate(case_ws.iter_rows(min_row=2, max_col=max_col, values_only=True), start=2):
            case_col = self.col_to_index(self.case_id_col.get())
            req_col = self.col_to_index(self.case_req_col.get())
            
            # 处理用例ID，避免None显示为'None'字符串
            if len(row) > case_col and row[case_col] is not None:
                case_val = str(row[case_col]).strip()
                if case_val == 'None':
                    case_val = ''
            else:
                case_val = ''
            
            # 处理需求ID，避免None显示为'None'字符串
            if len(row) > req_col and row[req_col] is not None:
                req_val = str(row[req_col]).strip()
                if req_val == 'None':
                    req_val = ''
            else:
                req_val = ''
            
            ws2.cell(row=idx, column=1, value=case_val)
            ws2.cell(row=idx, column=2, value=req_val)
            ws2.cell(row=idx, column=1).border = thin_border
            ws2.cell(row=idx, column=2).border = thin_border
        
        ws2.column_dimensions['A'].width = 35
        ws2.column_dimensions['B'].width = 35
        
        # 异常分析
        ws3 = wb.create_sheet("异常分析")
        ws3['A1'] = '异常类型'
        ws3['B1'] = 'ID'
        ws3['C1'] = '说明'
        for cell in [ws3['A1'], ws3['B1'], ws3['C1']]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
        
        row_idx = 2
        orphan_cases = [c for c, r in case_to_reqs_raw.items() if not r]
        orphan_reqs = [r for r, c in req_to_cases.items() if not c]
        
        for case_id in orphan_cases:
            ws3.cell(row=row_idx, column=1, value="孤儿用例")
            ws3.cell(row=row_idx, column=2, value=case_id)
            ws3.cell(row=row_idx, column=3, value="该用例未关联任何需求")
            for c in [1, 2, 3]:
                ws3.cell(row=row_idx, column=c).border = thin_border
            row_idx += 1
        
        for req_id in orphan_reqs:
            ws3.cell(row=row_idx, column=1, value="孤儿需求")
            ws3.cell(row=row_idx, column=2, value=req_id)
            ws3.cell(row=row_idx, column=3, value="该需求在用例文档中被引用，但无实际关联用例")
            for c in [1, 2, 3]:
                ws3.cell(row=row_idx, column=c).border = thin_border
            row_idx += 1
        
        ws3.column_dimensions['A'].width = 20
        ws3.column_dimensions['B'].width = 40
        ws3.column_dimensions['C'].width = 40
        
        # 需求缺失用例
        ws4 = wb.create_sheet("需求缺失用例")
        ws4['A1'] = '需求ID'
        ws4['B1'] = '状态'
        ws4['C1'] = '说明'
        for cell in [ws4['A1'], ws4['B1'], ws4['C1']]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
        
        row_idx = 2
        for req_id in sorted(missing_case_reqs):
            ws4.cell(row=row_idx, column=1, value=req_id)
            ws4.cell(row=row_idx, column=2, value="未覆盖")
            ws4.cell(row=row_idx, column=3, value="该需求在需求文档中存在，但未在测试用例中体现")
            for c in [1, 2, 3]:
                ws4.cell(row=row_idx, column=c).border = thin_border
            row_idx += 1
        
        ws4.column_dimensions['A'].width = 40
        ws4.column_dimensions['B'].width = 15
        ws4.column_dimensions['C'].width = 50
        
        # 统计汇总
        ws5 = wb.create_sheet("统计汇总")
        ws5['A1'] = '统计项'
        ws5['B1'] = '数量'
        ws5['C1'] = '说明'
        for cell in [ws5['A1'], ws5['B1'], ws5['C1']]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
        
        total_cases = len(case_to_reqs_raw)
        linked_cases = len([c for c, r in case_to_reqs_raw.items() if r])
        orphan_case_count = len(orphan_cases)
        total_reqs = len(req_to_cases)
        linked_reqs = len([r for r, c in req_to_cases.items() if c])
        orphan_req_count = len(orphan_reqs)
        coverage = (linked_reqs / len(all_req_ids_raw) * 100) if all_req_ids_raw else 0
        
        stats = [
            ('总用例数', total_cases, '用例文档中用例ID非空的总数'),
            ('双向关联用例数', linked_cases, '有关联需求的用例数'),
            ('孤儿用例数', orphan_case_count, '未关联任何需求的用例数'),
            ('', '', ''),
            ('用例中引用需求数', total_reqs, '用例文档中引用的需求ID总数'),
            ('双向关联需求数', linked_reqs, '有关联用例的需求数'),
            ('孤儿需求数', orphan_req_count, '有用例ID为空但引用的需求数'),
            ('', '', ''),
            ('需求文档总需求数', len(all_req_ids_raw), '需求文档中的需求总数'),
            ('需求缺失用例数', len(missing_case_reqs), '需求文档中有但用例未覆盖的需求数'),
            ('需求覆盖率', f'{coverage:.1f}%', '已覆盖需求数/需求文档总需求数'),
        ]
        
        row_idx = 2
        for item, count, desc in stats:
            ws5.cell(row=row_idx, column=1, value=item)
            ws5.cell(row=row_idx, column=2, value=count)
            ws5.cell(row=row_idx, column=3, value=desc)
            for c in [1, 2, 3]:
                ws5.cell(row=row_idx, column=c).border = thin_border
            row_idx += 1
        
        ws5.column_dimensions['A'].width = 25
        ws5.column_dimensions['B'].width = 15
        ws5.column_dimensions['C'].width = 50
        
        wb.save(output_path)
        print(f"溯源分析完成，结果已保存至：{output_path}")


def main():
    root = tk.Tk()
    app = TraceabilityTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
