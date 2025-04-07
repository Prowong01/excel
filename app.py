import os
import sys
from flask import Flask, render_template, request, send_file, jsonify
import webview
import threading
import tempfile
from main import process_excel
from compare_excel import compare_excel_files  # 添加导入
import pandas as pd
from datetime import datetime

def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
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

# 使用应用程序特定的临时目录
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'excel_processor_uploads')
PROCESSED_FOLDER = os.path.join(tempfile.gettempdir(), 'excel_processor_processed')
COMPARE_FOLDER = os.path.join(tempfile.gettempdir(), 'excel_processor_compare')  # 新增比较文件夹

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
        # Clear temporary folders
        for folder in [UPLOAD_FOLDER, PROCESSED_FOLDER]:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Error removing file {file_path}: {e}")

        # Save uploaded files
        saved_files = []
        for file in files:
            if file.filename:
                filepath = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(filepath)
                saved_files.append(filepath)

        # Process files
        all_data = []
        for file_path in saved_files:
            processed_df = process_excel(file_path)
            if processed_df is not None:
                all_data.append(processed_df)

        if all_data:
            # Merge all data
            final_df = pd.concat(all_data, ignore_index=True)
            
            # Save processed file
            output_path = os.path.join(PROCESSED_FOLDER, 'processed_data.xlsx')
            final_df.to_excel(output_path, index=False)
            
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
        return send_file(
            output_path,
            as_attachment=True,
            download_name='processed_data.xlsx'
        )
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
                os.remove(file_path)
            except Exception as e:
                print(f"清理文件失败 {file_path}: {e}")

        # 保存上传的文件
        file1_path = os.path.join(COMPARE_FOLDER, 'file1.xlsx')
        file2_path = os.path.join(COMPARE_FOLDER, 'file2.xlsx')
        
        file1.save(file1_path)
        file2.save(file2_path)

        # 执行比较
        merged_df, calc_df = compare_excel_files(file1_path, file2_path)
        
        if merged_df is not None and calc_df is not None:
            # 生成输出文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(COMPARE_FOLDER, f"compared_data_{timestamp}.xlsx")
            
            # 保存结果
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
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename
        )
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
