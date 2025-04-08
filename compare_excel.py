import pandas as pd
import os
from datetime import datetime
import traceback
import sys

def compare_excel_files(file1_path, file2_path):
    try:
        # 使用 utf-8 编码读取 Excel 文件
        df1 = pd.read_excel(file1_path, engine='openpyxl')  # 旧文件
        df2 = pd.read_excel(file2_path, engine='openpyxl')  # 新文件
        
        print(f"文件1(旧)行数: {len(df1)}")
        print(f"文件2(新)行数: {len(df2)}")
        
        # 确保两个DataFrame都有必要的列
        required_cols = ['post_id', 'video_views']
        for col in required_cols:
            if col not in df1.columns or col not in df2.columns:
                raise ValueError(f"缺少必要的列: {col}")
        
        # 获取两个DataFrame中的所有数值列
        numeric_cols = ['video_views']
        available_numeric_cols = [col for col in numeric_cols if col in df1.columns and col in df2.columns]
        
        # 将df1(旧文件)的相关列重命名，以避免合并时的冲突
        rename_dict = {col: f'{col}_old' for col in available_numeric_cols}
        df1 = df1.rename(columns=rename_dict)
        
        # 准备要合并的列
        merge_cols = ['post_id'] + [f'{col}_old' for col in available_numeric_cols]
        
        # 基于post_id合并两个DataFrame，以新文件(df2)为基准
        merged_df = pd.merge(
            df2,  # 新文件作为基准
            df1[merge_cols],  # 旧文件的数据
            on='post_id', 
            how='left'  # 保留所有新文件的记录
        )
        
        # 计算所有数值列的差异（新 - 旧）
        for col in available_numeric_cols:
            merged_df[f'{col}_difference'] = merged_df[col].fillna(0) - merged_df[f'{col}_old'].fillna(0)
        
        # 创建计算统计数据
        metrics = []
        values = []
        
        # 添加基础统计
        metrics.extend([
            'post_count (old)',
            'post_count (new)',
            'post_count_difference'
        ])
        values.extend([
            len(df1),
            len(df2),
            len(df2) - len(df1)
        ])
        
        # 添加每个数值列的统计
        for col in available_numeric_cols:
            metrics.extend([
                f'total_{col} (old)',
                f'total_{col} (new)',
                f'total_{col}_difference'
            ])
            values.extend([
                df1[f'{col}_old'].fillna(0).sum(),
                df2[col].fillna(0).sum(),
                df2[col].fillna(0).sum() - df1[f'{col}_old'].fillna(0).sum()
            ])
        
        calc_data = {
            'Metrics': metrics,
            'Values': values
        }
        
        calc_df = pd.DataFrame(calc_data)
        
        # 生成输出文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(
            os.path.dirname(file1_path),
            f"compared_data_{timestamp}.xlsx"
        )
        
        # 使用ExcelWriter保存多个sheet
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            merged_df.to_excel(writer, sheet_name='Data', index=False)
            calc_df.to_excel(writer, sheet_name='Calculation', index=False)
        
        print(f"比较完成！输出文件：{output_path}")
        
        # 打印统计信息
        matched_count = len(merged_df[merged_df[f'{available_numeric_cols[0]}_old'].notna()])
        unmatched_count = len(merged_df) - matched_count
        print(f"\n统计信息:")
        print(f"匹配的post_id数量: {matched_count}")
        print(f"未匹配的post_id数量: {unmatched_count}")
        print(f"帖子数量变化: {len(df2) - len(df1):,}")
        
        return merged_df, calc_df
        
    except Exception as e:
        print(f"比较文件时发生错误: {str(e)}")
        traceback.print_exc()
        return None, None

def main():
    try:
        if len(sys.argv) != 3:
            print("使用方法: python compare_excel.py <旧文件路径> <新文件路径>")
            print("例如: python compare_excel.py old_data.xlsx new_data.xlsx")
            return
        
        file1_path = sys.argv[1]  # 较旧的文件
        file2_path = sys.argv[2]  # 较新的文件
        
        # 检查文件是否存在
        if not os.path.exists(file1_path):
            print(f"错误: 找不到文件 {file1_path}")
            return
            
        if not os.path.exists(file2_path):
            print(f"错误: 找不到文件 {file2_path}")
            return
        
        print(f"\n开始比较文件:")
        print(f"文件1 (较旧): {file1_path}")
        print(f"文件2 (较新): {file2_path}")
        
        compare_excel_files(file1_path, file2_path)
        
    except Exception as e:
        print(f"程序运行错误: {str(e)}")
        traceback.print_exc()
    
    # 添加暂停，让用户看到结果
    print("\n处理完成。按回车键退出...")
    input()

if __name__ == "__main__":
    main()
