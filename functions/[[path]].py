import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime

class ScoreAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("重庆市潼南区塘坝文昌学校成绩计算工具 - by袁华")
        self.root.geometry("900x700")
        self.root.resizable(False, False)
        
        self.file_path = None
        self.df = None
        self.scores_columns = {
            '语文': 'H',
            '数学': 'K',
            '英语': 'N',
            '科学': 'Q',
            '道法': 'T'
        }
        
        self.setup_ui()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(main_frame, text="学生成绩批量分析工具", 
                 font=('微软雅黑', 16, 'bold')).pack(pady=20)
        
        # 打开文件按钮
        self.open_file_btn = ttk.Button(
            main_frame, text="📂 打开Excel成绩文件", 
            command=self.open_file, style='TButton',
            width=30
        )
        ttk.Style().configure('TButton', font=('微软雅黑', 12))
        self.open_file_btn.pack(pady=15)
        
        # 文件状态
        self.file_status = ttk.Label(
            main_frame, text="未加载文件，请先点击上方按钮选择Excel文件",
            font=('微软雅黑', 10), foreground='#666666'
        )
        self.file_status.pack(pady=5)
        
        # 分割线
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=20)
        
        # 总分设置
        ttk.Label(main_frame, text="各科总分设置（可修改）", 
                 font=('微软雅黑', 12, 'bold')).pack(pady=10)
        entry_frame = ttk.Frame(main_frame)
        entry_frame.pack(pady=8)
        self.score_entries = {}
        for idx, subject in enumerate(self.scores_columns.keys()):
            ttk.Label(entry_frame, text=f"{subject}：", font=('微软雅黑', 10)).grid(
                row=0, column=idx*2, padx=3, pady=5
            )
            entry = ttk.Entry(entry_frame, width=8, font=('微软雅黑', 10))
            entry.insert(0, "100")
            entry.grid(row=0, column=idx*2+1, padx=3, pady=5)
            self.score_entries[subject] = entry
        ttk.Label(entry_frame, text="分", font=('微软雅黑', 10)).grid(
            row=0, column=len(self.scores_columns)*2, padx=3
        )
        
        # 分析按钮
        self.analyze_btn = ttk.Button(
            main_frame, text="🚀 开始成绩分析", 
            command=self.analyze_scores, width=30
        )
        self.analyze_btn.pack(pady=20)
        
        # 结果展示
        ttk.Label(main_frame, text="分析结果预览", font=('微软雅黑', 12, 'bold')).pack(pady=10, anchor=tk.W)
        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text = tk.Text(
            result_frame, height=12, font=('微软雅黑', 9),
            yscrollcommand=scrollbar.set, state='disabled'
        )
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=2)
        scrollbar.config(command=self.result_text.yview)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪 | 等待加载文件")
        self.status_bar = ttk.Label(
            main_frame, textvariable=self.status_var, 
            relief=tk.SUNKEN, anchor=tk.W, padding=5
        )
        self.status_bar.pack(fill=tk.X, pady=10)
    
    def open_file(self):
        """打开Excel成绩文件"""
        try:
            file_path = filedialog.askopenfilename(
                title="选择Excel成绩文件",
                filetypes=[("Excel文件", "*.xlsx"), ("旧版Excel", "*.xls")],
                initialdir=os.path.expanduser("~")
            )
            if not file_path:
                return
            if not os.path.exists(file_path) or not file_path.lower().endswith(('.xlsx', '.xls')):
                messagebox.showerror("错误", "请选择有效的Excel文件（.xlsx/.xls）")
                return
            
            self.df = pd.read_excel(file_path, header=None, skiprows=4, engine='openpyxl')
            self.df.columns = [chr(65 + i) for i in range(len(self.df.columns))]
            
            required_cols = ['B'] + list(self.scores_columns.values())
            missing_cols = [col for col in required_cols if col not in self.df.columns]
            if missing_cols:
                raise ValueError(f"缺少必要列：{', '.join(missing_cols)}\n请检查Excel文件格式！")
            
            self.file_path = file_path
            file_name = os.path.basename(file_path)
            self.file_status.config(
                text=f"已加载：{file_name} | 共{len(self.df)}条数据",
                foreground='#28a745'
            )
            self.status_var.set(f"就绪 | 已加载{file_name}，可开始分析")
            messagebox.showinfo("成功", f"Excel文件加载成功！\n共读取{len(self.df)}条学生数据")
        except Exception as e:
            messagebox.showerror("文件加载失败", f"失败原因：{str(e)}")
            self.file_status.config(text="加载失败，请重新选择文件", foreground='#dc3545')
            self.status_var.set("错误 | 文件加载失败")
    
    def analyze_scores(self):
        """核心分析逻辑，生成Excel结果"""
        if self.df is None:
            messagebox.showerror("提示", "请先点击【打开Excel成绩文件】按钮加载文件！")
            return
        
        # 校验总分
        try:
            full_scores = {}
            for subject, entry in self.score_entries.items():
                val = entry.get().strip()
                if not val:
                    raise ValueError(f"请填写{subject}的总分！")
                score = float(val)
                if score <= 0:
                    raise ValueError(f"{subject}总分必须大于0！")
                full_scores[subject] = score
        except ValueError as e:
            messagebox.showerror("输入错误", str(e))
            return
        
        try:
            self.status_var.set("分析中 | 正在处理成绩数据，请稍候...")
            self.root.update_idletasks()
            df = self.df.copy()
            total_students = len(df)
            results_text = []  # 界面预览文本
            excel_data = []    # Excel表格数据（年级+班级）

            # 分数列转数值，空值填0
            for col in self.scores_columns.values():
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df['总分'] = df[list(self.scores_columns.values())].sum(axis=1, skipna=True)

            # ---------------------- 年级整体统计 ----------------------
            results_text.append("="*80)
            results_text.append("                    年级整体成绩统计报告")
            results_text.append("="*80)
            # 年级统计行（Excel）
            grade_row = ['年级整体', total_students, '—']
            for subject, col in self.scores_columns.items():
                students_to_count = max(1, round(total_students * 0.95))
                top_scores = df[col].nlargest(students_to_count)
                avg_score = top_scores.mean() if not top_scores.empty else 0.0
                
                excellent_cutoff = full_scores[subject] * 0.8
                pass_cutoff = full_scores[subject] * 0.6
                fail_cutoff = full_scores[subject] * 0.4
                excellent_count = len(df[df[col] >= excellent_cutoff])
                pass_count = len(df[df[col] >= pass_cutoff])
                fail_count = len(df[df[col] < fail_cutoff])
                
                excellent_rate = (excellent_count / total_students * 100) if total_students > 0 else 0.0
                pass_rate = (pass_count / total_students * 100) if total_students > 0 else 0.0
                fail_rate = (fail_count / total_students * 100) if total_students > 0 else 0.0

                # 界面文本
                results_text.append(f"\n{subject}科目：")
                results_text.append(f"  年级平均分（前95%学生）：{avg_score:.2f} 分")
                results_text.append(f"  优生人数：{excellent_count} 人 | 优生率：{excellent_rate:.2f}%")
                results_text.append(f"  及格人数：{pass_count} 人 | 及格率：{pass_rate:.2f}%")
                results_text.append(f"  差生人数：{fail_count} 人 | 差生率：{fail_rate:.2f}%")
                # Excel行追加（平均分、优生率、及格率、差生率）
                grade_row.extend([round(avg_score,2), f"{excellent_rate:.2f}%", f"{pass_rate:.2f}%", f"{fail_rate:.2f}%"])

            excel_data.append(grade_row)
            results_text.append("\n" + "="*80)
            results_text.append("                    各班成绩详细统计报告")
            results_text.append("="*80)

            # ---------------------- 分班级统计 ----------------------
            if total_students > 0:
                classes = sorted(df['B'].dropna().unique())
                for class_name in classes:
                    class_df = df[df['B'] == class_name].copy()
                    class_total = len(class_df)
                    if class_total == 0:
                        continue
                    # 班级统计行（Excel）
                    class_row = [f'{class_name}', class_total, f"{(class_total/total_students*100):.1f}%"]
                    # 界面文本
                    results_text.append(f"\n【班级：{class_name}】（学生总数：{class_total} 人）")

                    for subject, col in self.scores_columns.items():
                        class_stu_count = max(1, round(class_total * 0.95))
                        class_top_scores = class_df[col].nlargest(class_stu_count)
                        class_avg = class_top_scores.mean() if not class_top_scores.empty else 0.0
                        
                        excellent_cutoff = full_scores[subject] * 0.8
                        pass_cutoff = full_scores[subject] * 0.6
                        class_excellent = len(class_df[class_df[col] >= excellent_cutoff])
                        class_pass = len(class_df[class_df[col] >= pass_cutoff])
                        class_fail = class_total - class_pass
                        
                        class_excellent_rate = (class_excellent / class_total * 100) if class_total > 0 else 0.0
                        class_pass_rate = (class_pass / class_total * 100) if class_total > 0 else 0.0
                        class_fail_rate = (class_fail / class_total * 100) if class_total > 0 else 0.0

                        # 界面文本
                        results_text.append(f"  {subject}：")
                        results_text.append(f"    班级平均分：{class_avg:.2f} 分")
                        results_text.append(f"    优生：{class_excellent}人({class_excellent_rate:.2f}%) | 及格：{class_pass}人({class_pass_rate:.2f}%) | 差生：{class_fail}人({class_fail_rate:.2f}%)")
                        # Excel行追加
                        class_row.extend([round(class_avg,2), f"{class_excellent_rate:.2f}%", f"{class_pass_rate:.2f}%", f"{class_fail_rate:.2f}%"])

                    excel_data.append(class_row)
            else:
                results_text.append("\n暂无有效学生数据可统计")

            # ---------------------- 界面显示结果 ----------------------
            self.result_text.config(state='normal')
            self.result_text.delete('1.0', tk.END)
            self.result_text.insert('1.0', '\n'.join(results_text))
            self.result_text.config(state='disabled')

            # ---------------------- 生成Excel表格 ----------------------
            self.export_to_excel(excel_data, full_scores)

            self.status_var.set("分析完成 | 已生成Excel分析报告！")
            messagebox.showinfo("分析成功", f"成绩分析完成！\n✅ 界面显示结果预览\n✅ 已生成标准Excel分析报告（与原文件同目录）\n✅ 支持直接编辑/打印/二次统计")
        except Exception as e:
            messagebox.showerror("分析失败", f"失败原因：{str(e)}")
            self.status_var.set("分析失败 | 请检查文件格式或总分设置")
    
    def export_to_excel(self, excel_data, full_scores):
        """生成标准Excel分析报告，带表头、格式"""
        if not self.file_path or not excel_data:
            return
        try:
            # Excel表头构建（动态适配科目）
            header = ['统计维度', '学生总数', '年级占比']
            for subject in self.scores_columns.keys():
                header.extend([f'{subject}平均分', f'{subject}优生率', f'{subject}及格率', f'{subject}差生率'])
            
            # 构建DataFrame（Excel核心）
            df_excel = pd.DataFrame(excel_data, columns=header)
            # 导出路径：原Excel同目录，带时间戳（避免覆盖）
            output_dir = os.path.dirname(self.file_path)
            time_str = datetime.now().strftime('%Y%m%d_%H%M%S')
            excel_output = os.path.join(output_dir, f"成绩分析报告_{time_str}.xlsx")
            
            # 写入Excel并美化格式（调整列宽、居中）
            with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
                df_excel.to_excel(writer, sheet_name='成绩统计', index=False)
                # 获取工作表
                worksheet = writer.sheets['成绩统计']
                # 调整列宽（适配内容）
                for col in worksheet.columns:
                    max_length = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 20)
                    worksheet.column_dimensions[col_letter].width = adjusted_width
                # 所有内容居中对齐
                from openpyxl.styles import Alignment
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # 写入配置信息（新增工作表）
            config_data = [
                ['分析配置信息', ''],
                ['原数据文件', os.path.basename(self.file_path)],
                ['分析时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                ['统计规则', '1. 平均分取各班/年级前95%最高成绩；2. 优生≥80%总分，及格≥60%总分，差生<60%总分'],
                ['', ''],
                ['各科总分设置', ''],
            ] + [[subj, f'{score}分'] for subj, score in full_scores.items()]
            df_config = pd.DataFrame(config_data)
            df_config.to_excel(writer, sheet_name='分析配置', index=False, header=False)
            # 配置表列宽调整
            ws_config = writer.sheets['分析配置']
            ws_config.column_dimensions['A'].width = 15
            ws_config.column_dimensions['B'].width = 30

        except Exception as e:
            messagebox.showwarning("导出提示", f"Excel导出失败：{str(e)}\n💡 可手动复制界面结果，或检查是否安装openpyxl")
            print(f"Excel导出错误：{e}")

def main():
    """主函数：检查依赖+启动程序"""
    try:
        import openpyxl
    except ImportError:
        messagebox.showwarning("依赖缺失", "请先打开命令提示符，运行以下命令安装依赖：\npip install pandas openpyxl")
        return
    root = tk.Tk()
    ttk.Style().configure('.', font=('微软雅黑', 10))
    app = ScoreAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
