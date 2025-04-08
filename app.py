import os
import sys
from flask import Flask, render_template, request, send_file, jsonify
import webview
import threading
import tempfile
from main import process_excel
from compare_excel import compare_excel_files
import pandas as pd
from datetime import datetime

def get_resource_path(relative_path):
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

app = Flask(__name__,
    template_folder=get_resource_path('templates'),
    static_folder=get_resource_path('static')
)

# 使用系统临时目录
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'excel_processor_uploads')
PROCESSED_FOLDER = os.path.join(tempfile.gettempdir(), 'excel_processor_processed')
COMPARE_FOLDER = os.path.join(tempfile.gettempdir(), 'excel_processor_compare')

# 确保目录存在
for folder in [UPLOAD_FOLDER, PROCESSED_FOLDER, COMPARE_FOLDER]:
    os.makedirs(folder, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'files[]' not in request.files:
        return jsonify({'error': '没有选择文件'}), 400
    
    files = request.files.getlist('files[]')
    if not files:
        return jsonify({'error': '没有选择文件'}), 400

    try:
        # 清理临时文件夹
        for folder in [UPLOAD_FOLDER, PROCESSED_FOLDER]:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                except Exception as e:
                    print(f"Error removing file {file_path}: {e}")

        # 保存上传的文件
        saved_files = []
        for file in files:
            if file.filename:
                # 使用安全的文件名
                safe_filename = os.path.basename(file.filename)
                filepath = os.path.join(UPLOAD_FOLDER, safe_filename)
                file.save(filepath)
                saved_files.append(filepath)

        # 处理文件
        all_data = []
        for file_path in saved_files:
            processed_df = process_excel(file_path)
            if processed_df is not None:
                all_data.append(processed_df)

        if all_data:
            # 合并所有数据
            final_df = pd.concat(all_data, ignore_index=True)
            
            # 检查并处理重复的post_id
            if 'post_id' in final_df.columns:
                duplicate_mask = final_df['post_id'].duplicated(keep=False)
                duplicates = final_df[duplicate_mask]
                
                if not duplicates.empty:
                    print(f"发现{len(duplicates)}行重复的post_id，进行合并处理")
                    
                    # 定义需要合并的数值列
                    numeric_cols = ['video_views', 'like', 'comment', 'share', 'collect']
                    # 只使用存在的列
                    numeric_cols = [col for col in numeric_cols if col in final_df.columns]
                    
                    # 保留第一次出现的记录
                    unique_records = final_df[~final_df['post_id'].duplicated(keep='first')]
                    # 获取重复的记录（不包括第一次出现的）
                    duplicate_records = final_df[final_df['post_id'].duplicated(keep='first')]
                    
                    # 对重复记录按post_id分组并求和
                    if numeric_cols:
                        sums = duplicate_records.groupby('post_id')[numeric_cols].sum()
                        
                        # 更新原始数据中的数值
                        for post_id in sums.index:
                            mask = unique_records['post_id'] == post_id
                            for col in numeric_cols:
                                unique_records.loc[mask, col] += sums.loc[post_id, col]
                    
                    final_df = unique_records
                    print(f"合并后的数据行数: {len(final_df)}")
            
            # 保存处理后的文件
            output_path = os.path.join(PROCESSED_FOLDER, 'processed_data.xlsx')
            final_df.to_excel(output_path, index=False, engine='openpyxl')
            
            return jsonify({
                'success': True,
                'message': '文件处理成功',
                'download_ready': True
            })
        else:
            return jsonify({
                'error': '文件处理失败'
            }), 500

    except Exception as e:
        return jsonify({
            'error': f'处理过程中出错: {str(e)}'
        }), 500

@app.route('/download')
def download():
    output_path = os.path.join(PROCESSED_FOLDER, 'processed_data.xlsx')
    if os.path.exists(output_path):
        try:
            return send_file(
                output_path,
                as_attachment=True,
                download_name='processed_data.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            return jsonify({'error': f'下载文件时出错: {str(e)}'}), 500
    return jsonify({'error': '文件不存在'}), 404

@app.route('/compare', methods=['POST'])
def compare_files():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': '需要上传两个文件'}), 400
    
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    if not file1 or not file2:
        return jsonify({'error': '请选择两个文件'}), 400

    try:
        # 清理比较文件夹
        for filename in os.listdir(COMPARE_FOLDER):
            file_path = os.path.join(COMPARE_FOLDER, filename)
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as e:
                print(f"清理文件失败 {file_path}: {e}")

        # 保存上传的文件，使用安全的文件名
        file1_path = os.path.join(COMPARE_FOLDER, 'file1.xlsx')
        file2_path = os.path.join(COMPARE_FOLDER, 'file2.xlsx')
        
        file1.save(file1_path)
        file2.save(file2_path)

        # 执行比较
        merged_df, calc_df = compare_excel_files(file1_path, file2_path)
        
        if merged_df is not None and calc_df is not None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(COMPARE_FOLDER, f"compared_data_{timestamp}.xlsx")
            
            # 使用 openpyxl 引擎保存 Excel
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                merged_df.to_excel(writer, sheet_name='Data', index=False)
                calc_df.to_excel(writer, sheet_name='Calculation', index=False)
            
            return jsonify({
                'success': True,
                'message': '文件比较成功',
                'download_ready': True,
                'timestamp': timestamp
            })
        else:
            return jsonify({
                'error': '文件比较失败'
            }), 500

    except Exception as e:
        return jsonify({
            'error': f'比较过程中出错: {str(e)}'
        }), 500

@app.route('/download_comparison/<timestamp>')
def download_comparison(timestamp):
    filename = f"compared_data_{timestamp}.xlsx"
    file_path = os.path.join(COMPARE_FOLDER, filename)
    
    if os.path.exists(file_path):
        try:
            return send_file(
                file_path,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            return jsonify({'error': f'下载文件时出错: {str(e)}'}), 500
    return jsonify({'error': '比较结果文件不存在'}), 404

def start_server():
    app.run(port=5000, debug=False)

def main():
    t = threading.Thread(target=start_server)
    t.daemon = True
    t.start()
    
    webview.create_window(
        'Excel 处理工具', 
        'http://localhost:5000',
        width=800,
        height=600,
        resizable=True
    )
    webview.start()

if __name__ == '__main__':
    main()
