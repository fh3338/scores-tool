# -*- coding: utf-8 -*-
"""
重庆市潼南区塘坝文昌学校成绩计算工具 - Web版
功能：Excel上传、年级/班级成绩统计、Excel报告导出
统计规则：
1.  平均分取各班/年级前95%最高成绩
2.  优生 ≥ 80% 单科总分
3.  及格 ≥ 60% 单科总分
4.  差生 < 40% 单科总分（已修正）
GitHub托管专用，无本地文件依赖，可直接克隆运行
"""

# 导入必要依赖（均为PyPI公开库，GitHub克隆后可通过requirements.txt安装）
from flask import Flask, request, jsonify, send_file
import pandas as pd
from datetime import datetime
from openpyxl.styles import Alignment
import io

# 1. 初始化Flask应用（符合Web服务规范，无硬编码）
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件16M，避免超大文件攻击

# 2. 成绩分析核心类（无本地文件操作，适配GitHub托管环境）
class ScoreAnalyzer:
    def __init__(self):
        """初始化核心配置，与原GUI工具保持一致"""
        self.scores_columns = {
            '语文': 'H',
            '数学': 'K',
            '英语': 'N',
            '科学': 'Q',
            '道法': 'T'
        }
        self.df = None  # 存储Excel解析数据
        self.excel_buffer = io.BytesIO()  # 内存缓冲区存储Excel报告，无本地文件生成
        self.analysis_result = ""  # 存储文本格式分析结果

    def load_excel_file(self, file_stream):
        """
        加载上传的Excel文件（适配Web文件流，无本地路径依赖）
        :param file_stream: Flask上传的文件二进制流
        :return: (是否成功, 提示信息)
        """
        try:
            # 保持原Excel解析规则：跳过前4行、列名用A/B/C...命名、支持.xlsx格式
            self.df = pd.read_excel(
                file_stream,
                header=None,
                skiprows=4,
                engine='openpyxl'
            )
            self.df.columns = [chr(65 + i) for i in range(len(self.df.columns))]
            
            # 校验必要列（避免无效Excel文件）
            required_cols = ['B'] + list(self.scores_columns.values())
            missing_cols = [col for col in required_cols if col not in self.df.columns]
            if missing_cols:
                raise ValueError(f"缺少必要数据列：{', '.join(missing_cols)}，请检查Excel格式！")
            
            if len(self.df) == 0:
                raise ValueError("Excel文件中无有效学生成绩数据！")
            
            return True, f"文件加载成功，共读取{len(self.df)}条学生记录"
        except Exception as e:
            return False, f"文件加载失败：{str(e)}"

    def analyze_scores(self, full_scores):
        """
        核心成绩统计（修正差生判定规则：<40%总分）
        :param full_scores: 各科总分配置字典
        :return: (是否成功, 提示信息)
        """
        if self.df is None:
            return False, "请先加载有效的Excel成绩文件！"
        
        try:
            df = self.df.copy()
            total_students = len(df)
            results_text = []
            excel_data = []

            # 分数预处理：转数值类型、空值填充为0
            for col in self.scores_columns.values():
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df['总分'] = df[list(self.scores_columns.values())].sum(axis=1, skipna=True)

            # ---------------------- 年级整体统计 ----------------------
            results_text.append("=" * 80)
            results_text.append("                    年级整体成绩统计报告")
            results_text.append("=" * 80)
            grade_row = ['年级整体', total_students, '—']
            
            for subject, col in self.scores_columns.items():
                # 前95%学生平均分（剔除末尾5%异常低分，保持原逻辑）
                students_to_count = max(1, round(total_students * 0.95))
                top_scores = df[col].nlargest(students_to_count)
                avg_score = top_scores.mean() if not top_scores.empty else 0.0
                
                # 判定阈值（修正差生阈值为40%总分）
                excellent_cutoff = full_scores[subject] * 0.8  # 优生≥80%
                pass_cutoff = full_scores[subject] * 0.6       # 及格≥60%
                fail_cutoff = full_scores[subject] * 0.4       # 差生<40%（已修正核心规则）
                
                # 统计各类学生数量
                excellent_count = len(df[df[col] >= excellent_cutoff])
                pass_count = len(df[df[col] >= pass_cutoff])
                fail_count = len(df[df[col] < fail_cutoff])    # 差生统计使用修正后的阈值
                
                # 率值计算（避免除零错误）
                excellent_rate = (excellent_count / total_students * 100) if total_students > 0 else 0.0
                pass_rate = (pass_count / total_students * 100) if total_students > 0 else 0.0
                fail_rate = (fail_count / total_students * 100) if total_students > 0 else 0.0

                # 整理文本结果
                results_text.append(f"\n{subject}科目：")
                results_text.append(f"  年级平均分（前95%学生）：{avg_score:.2f} 分")
                results_text.append(f"  优生人数：{excellent_count} 人 | 优生率：{excellent_rate:.2f}%")
                results_text.append(f"  及格人数：{pass_count} 人 | 及格率：{pass_rate:.2f}%")
                results_text.append(f"  差生人数：{fail_count} 人 | 差生率：{fail_rate:.2f}%")  # 对应修正后的规则
                
                # 整理Excel报表数据
                grade_row.extend([
                    round(avg_score, 2),
                    f"{excellent_rate:.2f}%",
                    f"{pass_rate:.2f}%",
                    f"{fail_rate:.2f}%"
                ])

            excel_data.append(grade_row)
            results_text.append("\n" + "=" * 80)
            results_text.append("                    各班成绩详细统计报告")
            results_text.append("=" * 80)

            # ---------------------- 分班级统计 ----------------------
            if total_students > 0:
                classes = sorted(df['B'].dropna().unique())
                for class_name in classes:
                    class_df = df[df['B'] == class_name].copy()
                    class_total = len(class_df)
                    if class_total == 0:
                        continue
                    
                    class_row = [f'{class_name}', class_total, f"{(class_total/total_students*100):.1f}%"]
                    results_text.append(f"\n【班级：{class_name}】（学生总数：{class_total} 人）")

                    for subject, col in self.scores_columns.items():
                        # 班级前95%平均分
                        class_stu_count = max(1, round(class_total * 0.95))
                        class_top_scores = class_df[col].nlargest(class_stu_count)
                        class_avg = class_top_scores.mean() if not class_top_scores.empty else 0.0
                        
                        # 判定阈值（同步修正差生阈值为40%）
                        excellent_cutoff = full_scores[subject] * 0.8
                        pass_cutoff = full_scores[subject] * 0.6
                        fail_cutoff = full_scores[subject] * 0.4
                        
                        # 班级统计
                        class_excellent = len(class_df[class_df[col] >= excellent_cutoff])
                        class_pass = len(class_df[class_df[col] >= pass_cutoff])
                        class_fail = len(class_df[class_df[col] < fail_cutoff])  # 同步修正
                        
                        # 班级率值计算
                        class_excellent_rate = (class_excellent / class_total * 100) if class_total > 0 else 0.0
                        class_pass_rate = (class_pass / class_total * 100) if class_total > 0 else 0.0
                        class_fail_rate = (class_fail / class_total * 100) if class_total > 0 else 0.0

                        # 整理班级文本结果
                        results_text.append(f"  {subject}：")
                        results_text.append(f"    班级平均分：{class_avg:.2f} 分")
                        results_text.append(f"    优生：{class_excellent}人({class_excellent_rate:.2f}%) | 及格：{class_pass}人({class_pass_rate:.2f}%) | 差生：{class_fail}人({class_fail_rate:.2f}%)")
                        
                        # 整理班级Excel数据
                        class_row.extend([
                            round(class_avg, 2),
                            f"{class_excellent_rate:.2f}%",
                            f"{class_pass_rate:.2f}%",
                            f"{class_fail_rate:.2f}%"
                        ])

                    excel_data.append(class_row)
            else:
                results_text.append("\n暂无有效学生数据可进行统计分析")

            # 保存结果到实例属性
            self.analysis_result = '\n'.join(results_text)
            # 生成Excel分析报告（内存缓冲区，无本地文件）
            self._generate_excel_report(excel_data, full_scores)
            
            return True, "成绩分析完成，已生成Excel格式分析报告"
        except Exception as e:
            return False, f"成绩分析失败：{str(e)}"

    def _generate_excel_report(self, excel_data, full_scores):
        """
        生成Excel分析报告（内存缓冲区，适配GitHub无本地写入权限环境）
        :param excel_data: 统计数据列表
        :param full_scores: 各科总分配置
        """
        if not excel_data:
            return
        
        # 构建Excel表头
        header = ['统计维度', '学生总数', '年级占比']
        for subject in self.scores_columns.keys():
            header.extend([
                f'{subject}平均分',
                f'{subject}优生率',
                f'{subject}及格率',
                f'{subject}差生率'  # 表头同步更新，对应修正后的规则
            ])
        
        # 构建成绩统计DataFrame
        df_excel = pd.DataFrame(excel_data, columns=header)
        
        # 写入内存Excel缓冲区
        with pd.ExcelWriter(self.excel_buffer, engine='openpyxl') as writer:
            # 工作表1：成绩统计（格式化，支持直接打印）
            df_excel.to_excel(writer, sheet_name='成绩统计', index=False)
            worksheet = writer.sheets['成绩统计']
            
            # 调整列宽（适配内容显示）
            for col in worksheet.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except Exception:
                        pass
                adjusted_width = min(max_length + 2, 20)
                worksheet.column_dimensions[col_letter].width = adjusted_width
            
            # 所有内容居中对齐
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # 工作表2：分析配置说明（修正差生规则注释）
            config_data = [
                ['分析配置信息', ''],
                ['分析时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                ['统计规则', '1. 平均分取各班/年级前95%最高成绩；2. 优生≥80%总分；3. 及格≥60%总分；4. 差生<40%总分（已修正）'],
                ['', ''],
                ['各科总分设置', ''],
            ] + [[subj, f'{score}分'] for subj, score in full_scores.items()]
            
            df_config = pd.DataFrame(config_data)
            df_config.to_excel(writer, sheet_name='分析配置', index=False, header=False)
            
            # 配置表列宽调整
            ws_config = writer.sheets['分析配置']
            ws_config.column_dimensions['A'].width = 15
            ws_config.column_dimensions['B'].width = 30
        
        # 重置缓冲区指针（关键：确保下载时能读取到完整内容）
        self.excel_buffer.seek(0)

# 3. Flask API接口（RESTful风格，适配GitHub托管后的Web访问）
@app.route('/', methods=['GET'])
def health_check():
    """健康检查接口，用于验证服务是否正常运行"""
    return jsonify({
        "code": 200,
        "msg": "成绩分析服务已正常启动（GitHub托管版）",
        "api_doc": {
            "endpoint": "/analyze",
            "method": "POST",
            "params": {
                "file": "必填，Excel成绩文件（.xlsx格式）",
                "chinese": "可选，语文总分（默认100）",
                "math": "可选，数学总分（默认100）",
                "english": "可选，英语总分（默认100）",
                "science": "可选，科学总分（默认100）",
                "politics": "可选，道法总分（默认100）"
            },
            "return": "Excel格式成绩分析报告"
        }
    }), 200

@app.route('/analyze', methods=['POST'])
def analyze_api():
    """核心分析接口：接收Excel上传，返回分析报告"""
    try:
        # 1. 校验上传文件
        if 'file' not in request.files:
            return jsonify({"code": 400, "msg": "未上传任何Excel文件"}), 400
        
        file = request.files['file']
        if file.filename == '' or not file.filename.lower().endswith('.xlsx'):
            return jsonify({"code": 400, "msg": "请上传有效的.xlsx格式Excel文件"}), 400
        
        # 2. 接收各科总分配置（默认100分，支持自定义）
        full_scores = {
            '语文': float(request.form.get('chinese', 100)),
            '数学': float(request.form.get('math', 100)),
            '英语': float(request.form.get('english', 100)),
            '科学': float(request.form.get('science', 100)),
            '道法': float(request.form.get('politics', 100))
        }
        
        # 3. 校验总分配置有效性
        for subj, score in full_scores.items():
            if score <= 0:
                return jsonify({"code": 400, "msg": f"{subj}总分必须大于0"}), 400
        
        # 4. 执行成绩分析
        analyzer = ScoreAnalyzer()
        load_success, load_msg = analyzer.load_excel_file(file.stream)
        if not load_success:
            return jsonify({"code": 500, "msg": load_msg}), 500
        
        analyze_success, analyze_msg = analyzer.analyze_scores(full_scores)
        if not analyze_success:
            return jsonify({"code": 500, "msg": analyze_msg}), 500
        
        # 5. 返回Excel文件下载（带时间戳，避免文件名重复）
        return send_file(
            analyzer.excel_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"成绩分析报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
    except Exception as e:
        return jsonify({"code": 500, "msg": f"服务器内部错误：{str(e)}"}), 500

# 4. 启动服务（适配本地调试与GitHub托管环境）
if __name__ == "__main__":
    app.run(debug=False, host='0.0.0.0', port=5000)
