from flask import Flask, request, jsonify
import os
from datetime import datetime
app = Flask(__name__)
class ScoreAnalyzer:   
    
            for col in self.scores_columns.values():
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df['总分'] = df[list(self.scores_columns.values())].sum(axis=1, skipna=True)
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
                @app.route('/analyze', methods=['POST'])
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
