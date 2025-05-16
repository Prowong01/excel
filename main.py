import os
import pandas as pd
from datetime import datetime
import traceback
import re
import sys
import numpy as np

# 定义主列模板
MASTER_COLUMNS = [
    "post_id", "post", "network", "profile", "domestic_overseas_label", "published_date", "date", # <--在此处添加 "date"
    "video_views", "playthrough_rate", "avg_play_duration", "video_link",
    "like", "comment", "share", "collect", "subscribers"
]

# 平台名称映射（基于文件名关键词）
PLATFORM_MAPPING = {
    "抖音": "抖音",
    "小红书": "小红书",
    "快手": "快手",
    "哔哩哔哩": "哔哩哔哩",
    "视频号": "视频号",
    "微信视频号": "微信视频号",
    "B站": "B站",
    "X": "X",
    "Youtube": "Youtube",
    "Facebook": "Facebook",
    "TikTok": "TikTok"
}

# 增强版列名映射
COLUMN_MAPPING = {
    # 中文映射
    "视频id": "post_id",
    "作品名称": "post",
    "作品": "post",
    "笔记标题": "post",
    "视频标题": "post",
    "视频描述": "post",
    "账号": "profile",
    "发布时间": "published_date",
    "首次发布时间": "published_date",
    "播放量": "video_views",
    "观看量": "video_views",
    "阅读（播放）": "video_views",
    "完播率": "playthrough_rate",
    "平均播放时长": "avg_play_duration",
    "人均观看时长": "avg_play_duration",
    "点赞量": "like",
    "点赞": "like",
    "喜欢量": "like",
    "喜欢": "like",
    "评论量": "comment",
    "评论": "comment",
    "分享量": "share",
    "分享": "share",
    "推荐": "share",
    "收藏量": "collect",
    "收藏": "collect",
    "粉丝增量": "subscribers",
    "涨粉量": "subscribers",
    "涨粉": "subscribers",
    "关注量": "subscribers",

    # 英文映射
    "post_id": "post_id",
    "subscribers_gained_from_video": "subscribers",
    "video_views": "video_views",
    "average_video_time_watched_seconds_": "avg_play_duration",
    "full_video_view_rate": "playthrough_rate",
    "likes": "like",
    "comments": "comment",
    "shares": "share",
    "date": "published_date",
    "profile": "profile",
    "network": "network",
    "post": "post",
    "link": "video_link",
    "video_link": "video_link",
    "video_url": "video_link",
    "url": "video_link",
}

# 修改这个函数来获取正确的应用程序路径
def get_application_path():
    # 检查是否是打包的可执行文件
    if getattr(sys, 'frozen', False):
        # 如果是 PyInstaller 打包的可执行文件
        application_path = os.path.dirname(sys.executable)
    else:
        # 如果是普通 Python 脚本
        application_path = os.path.dirname(os.path.abspath(__file__))
    
    return application_path

def is_foreign_file(filename):
    """判断文件是否为国外数据文件"""
    return "国外" in filename

def determine_game_label(post):
    """根据post内容判断游戏标签"""
    if not isinstance(post, str):
        return "others"
    
    post = post.lower()  # 转换为小写以便统一匹配
    
    # 使用字典存储游戏关键词及对应的标签
    game_rules = {
        "poe2": ["poe", "流放之路", "流放", "pathofexile", "path of exile"],
        "PUBG": ["pubg", "吃鸡", "绝地求生", "和平精英","游戏小剧场"],
        "Warframe": ["warframe", "星际战甲"],
        "P2": ["exoborne"],
        "Dyinglight": ["消逝的光芒", "消光", "拉万", "lawan", "克兰", "clane", "消失的光芒", "dyinglight", "dying light"],
        "Nikke": ["nikke", "妮姬", "胜利女神"],
        "英雄联盟": ["leagueoflegends", "lol", "英雄联盟"],
        "王者荣耀": ["王者荣耀", "hok", "honor of kings", "honorofkings"],
        "暗区突围": ["暗区突围"],
        "TGA": ["tga"],
        "堡垒之夜": ["堡垒之夜", "fortnite"],
        "Zenless": ["zenless"],
        "Roblox": ["roblox", "罗布乐思"],
        "无尽对决": ["无尽对决", "mlbb"],
        "绝区零": ["绝区零"],
        "Diablo": ["diablo", "暗黑破坏神", "暗黑"],
        "美国大选": ["美国大选"],
        "黑神话悟空": ["黑神话", "黑猴"],
        "我的世界": ["我的世界"],
        "Dune": ["沙丘"],
        "月光骑士": ["月光骑士"],
        "原子之心": ["原子之心"],
        "漫威争锋": ["漫威争锋", "marvel"],
        "怪物猎人": ["怪物猎人", "monsterhunter"]
    }
    
    # 遍历规则字典
    for label, keywords in game_rules.items():
        if any(keyword in post for keyword in keywords):
            return label
    
    return "others"

def process_excel(file_path):
    """
    处理单个Excel或CSV文件，进行数据清洗、转换和标准化。
    支持表头在非第一行的情况。

    Args:
        file_path (str): 输入文件的路径。

    Returns:
        pandas.DataFrame or None: 处理后的DataFrame，如果处理失败则返回None。
    """
    try:
        filename = os.path.basename(file_path)
        is_foreign = is_foreign_file(filename)
        
        # 读取前几行来确定表头位置
        if file_path.endswith('.xlsx'):
            # 先不指定header，读取前5行
            df_head = pd.read_excel(file_path, nrows=5, header=None)
        else:
            encodings = ['utf-8', 'gbk', 'gb2312', 'utf-16']
            df_head = None
            for encoding in encodings:
                try:
                    df_head = pd.read_csv(file_path, nrows=5, header=None, encoding=encoding)
                    print(f"成功使用 {encoding} 编码读取CSV文件")
                    break
                except Exception as e:
                    continue
            if df_head is None:
                raise Exception("无法使用任何编码方式读取CSV文件")

        # 确定真正的表头行
        header_row = 0
        max_matches = 0
        
        # 检查前5行，看哪一行与COLUMN_MAPPING的键匹配度最高
        for i in range(min(5, len(df_head))):
            # 将该行的值转换为字符串并清理
            row_values = [str(x).strip() if pd.notna(x) else '' for x in df_head.iloc[i]]
            current_matches = sum(1 for col in row_values if col in COLUMN_MAPPING)
            print(f"第 {i+1} 行匹配的列数: {current_matches}")
            if current_matches > max_matches:
                max_matches = current_matches
                header_row = i

        print(f"使用第 {header_row + 1} 行作为表头")
        
        # 使用确定的表头行重新读取文件
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, header=header_row)
        else:
            df = pd.read_csv(file_path, encoding=encoding, header=header_row)
        
        print(f"成功读取文件，行数: {len(df)}")
        
        # 打印列名用于调试
        
        # 清理列名
        df.columns = (
            df.columns.str.strip()
            .str.lower()  # 统一转为小写
            .str.replace(r'[^\w]+', '_', regex=True)  # 替换所有非单词字符为下划线
            .str.replace(r'_+', '_', regex=True)  # 合并连续下划线
        )
        
        # 映射列名
        df.rename(columns=lambda x: COLUMN_MAPPING.get(x, x), inplace=True)
        print(f"映射后列名: {df.columns.tolist()}")

        # 处理重复列
        if len(df.columns) != len(set(df.columns)):
            print("警告: 检测到重复列名")
            seen = {}
            duplicates = set()
            
            # 遍历列名并标记重复列
            for i, col in enumerate(df.columns):
                if col in seen:
                    print(f"处理重复列 '{col}'，位置 {i}")
                    duplicates.add(i)
                    # 合并数据到第一个出现的列
                    first_idx = seen[col]
                    # 数值列相加，非数值列保留第一个非空值
                    if df.iloc[:, first_idx].dtype in [np.int64, np.float64]:
                        df.iloc[:, first_idx] += pd.to_numeric(df.iloc[:, i], errors='coerce').fillna(0)
                    else:
                        df.iloc[:, first_idx] = df.iloc[:, first_idx].combine(
                            df.iloc[:, i], 
                            lambda x, y: x if pd.notna(x) else y
                        )
                else:
                    seen[col] = i
            
            # 删除重复列
            df = df.drop(df.columns[list(duplicates)], axis=1)
            print(f"去重后列名: {df.columns.tolist()}")
        
        # 检查并添加必要的列
        for col in ['post', 'profile', 'published_date']:
            if col not in df.columns:
                print(f"警告: 缺少关键列 '{col}'")
                if col == 'profile':
                    # 尝试从文件名中提取 profile
                    try:
                        # 移除文件扩展名
                        filename_without_ext = os.path.splitext(filename)[0]
                        # 查找并提取 profile
                        for platform in PLATFORM_MAPPING.keys():
                            if platform in filename_without_ext:
                                # 找到平台名称在文件名中的位置
                                platform_index = filename_without_ext.find(platform)
                                # 从平台名称后的第一个'-'开始截取
                                remaining = filename_without_ext[platform_index + len(platform):]
                                if remaining.startswith('-'):
                                    # 去掉开头的'-'并获取剩余部分作为profile
                                    profile_name = remaining[1:].strip()
                                    df[col] = profile_name
                                    print(f"从文件名成功提取 profile: {profile_name}")
                                    break
                        else:
                            # 如果没有找到平台标识，直接尝试获取最后一部分
                            if '-' in filename_without_ext:
                                profile_name = filename_without_ext.split('-')[-1].strip()
                                df[col] = profile_name
                                print(f"从文件名提取 profile（无平台标识）: {profile_name}")
                            else:
                                df[col] = "未知" + col
                                print("无法从文件名提取 profile")
                    except Exception as e:
                        print(f"从文件名提取 profile 时出错: {str(e)}")
                        df[col] = "未知" + col
                else:
                    df[col] = "未知" + col  # 添加默认列
        
        # 先处理 published_date 格式
        if 'published_date' in df.columns:
            def convert_date(date_val):
                if pd.isna(date_val):
                    return date_val
                try:
                    # 转换为字符串
                    date_str = str(date_val).strip()
                    
                    # 处理中文格式日期
                    if '年' in date_str and '月' in date_str and '日' in date_str:
                        # 提取年月日
                        match = re.search(r'(\d{4})年(\d{2})月(\d{2})日\s*(\d{2})?:?(\d{2})?:?(\d{2})?', date_str)
                        if match:
                            year, month, day, hours, minutes, seconds = match.groups()
                            # 如果时分秒为None，设置为00
                            hours = hours or "00"
                            minutes = minutes or "00"
                            seconds = seconds or "00"
                            return f"{year}-{month}-{day}_{hours}:{minutes}:{seconds}"
                    
                    # 尝试用pandas处理其他格式
                    parsed_date = pd.to_datetime(date_str, errors='coerce')
                    if pd.notna(parsed_date):
                        return parsed_date.strftime('%Y-%m-%d_%H:%M:%S')
                    
                    return date_str
                except Exception as e:
                    print(f"日期转换错误 ({date_val}): {str(e)}")
                    return date_val

            df['published_date'] = df['published_date'].apply(convert_date)

        # 生成network字段
        if is_foreign:
            # 对于国外文件，直接使用数据中的network列
            if 'network' in df.columns:
                print("使用数据中的network列")
            else:
                df['network'] = "国外平台"
        else:
            # 对于国内文件，从文件名推断平台
            found_platform = None
            for platform_name in PLATFORM_MAPPING:
                if platform_name in filename:
                    found_platform = PLATFORM_MAPPING[platform_name]
                    break
            
            df['network'] = found_platform if found_platform else 'Unknown'
            print(f"从文件名推断平台: {df['network'].iloc[0]}")
        
        # 预处理：强制规范化post内容（同时处理国内外数据）
        def clean_post(text):
            if pd.isna(text): return ""
            return re.sub(r'\s+', ' ', str(text).strip()).replace('\n', ' ')

        df['post'] = df['post'].apply(clean_post)

        # 处理post_id        
        # 检查post_id是否存在（包括模糊匹配）
        post_id_cols = [c for c in df.columns if 'post_id' in c]
        if post_id_cols:
            # 如果找到post_id列，直接使用第一个匹配的列
            print(f"使用已有的post_id列: {post_id_cols[0]}")
            df['post_id'] = df[post_id_cols[0]].astype(str).str.strip()
        else:
            print("生成post_id")
            post_ids = []
            for idx, row in df.iterrows():
                try:
                    post = str(row.get('post', '')) if pd.notna(row.get('post', '')) else 'unknown_post'
                    # 替换post中的换行符，避免因换行导致的识别问题
                    post = post.replace('\n', ' ').replace('\r', '')
                    
                    network = str(row.get('network', '')) if pd.notna(row.get('network', '')) else 'unknown_network'
                    profile = str(row.get('profile', '')) if pd.notna(row.get('profile', '')) else 'unknown_profile'
                    
                    # 处理发布日期
                    pub_date = row.get('published_date', '')
                    if pd.isna(pub_date):
                        pub_date_str = 'unknown_date'
                    elif isinstance(pub_date, str):
                        pub_date_str = pub_date.replace('/', '-').replace(' ', '_')
                    else:
                        try:
                            pub_date_str = pd.Timestamp(pub_date).strftime('%Y-%m-%d_%H:%M:%S')
                        except:
                            pub_date_str = str(pub_date).replace('/', '-').replace(' ', '_')
                    
                    post_id = f"{post}_{network}_{profile}_{pub_date_str}"
                    post_ids.append(post_id)
                except Exception as e:
                    print(f"生成第{idx}行post_id时出错: {str(e)}")
                    post_ids.append(f"error_row_{idx}")
            
            df['post_id'] = post_ids

        # 设置国内/国外标签
        if is_foreign:
            # 如果是国外文件，直接全部标记为"国外"
            df['domestic_overseas_label'] = '国外'
            print("文件标记为国外数据")
        else:
            # 对于国内文件，使用现有逻辑判断
            domestic_labels = []
            for idx, profile in enumerate(df['profile']):
                try:
                    if pd.isna(profile):
                        domestic_labels.append('未知')
                    else:
                        profile_str = str(profile)
                        overseas_keywords = ['海外', '国际', 'Global']
                        is_overseas = any(kw in profile_str for kw in overseas_keywords)
                        domestic_labels.append('国外' if is_overseas else '国内')
                except Exception as e:
                    print(f"处理第{idx}行区域标签时出错: {str(e)}")
                    domestic_labels.append('未知')
            
            df['domestic_overseas_label'] = domestic_labels
            print("已生成国内/国外标签")
        
        # 处理 video_link 列
        # 修改处理 video_link 列的逻辑
        # 在 process_excel 函数中找到处理 video_link 的部分，替换为：
        # 处理 video_link 列
        possible_link_columns = ['link', 'video_link', 'video_url', 'url', 'ahmain']
        link_found = False
        for link_col in possible_link_columns:
            if link_col in df.columns:
                df['video_link'] = df[link_col]
                link_found = True
                print(f"使用 {link_col} 列作为 video_link")
                break
        
        if not link_found:
            df['video_link'] = pd.NA
            print("未找到任何链接列，video_link 设置为空")

        # 处理 published_date 格式
        if 'published_date' in df.columns:
            def convert_date(date_val):
                if pd.isna(date_val):
                    return date_val
                try:
                    # 转换为字符串
                    date_str = str(date_val).strip()
                    
                    # 处理中文格式日期
                    if '年' in date_str and '月' in date_str and '日' in date_str:
                        # 提取年月日
                        match = re.search(r'(\d{4})年(\d{2})月(\d{2})日', date_str)
                        if match:
                            year, month, day = match.groups()
                            return f"{year}-{month}-{day}"
                    
                    # 尝试用pandas处理其他格式
                    parsed_date = pd.to_datetime(date_str, errors='coerce')
                    if pd.notna(parsed_date):
                        return parsed_date.strftime('%Y-%m-%d')
                    
                    return date_str
                except Exception as e:
                    print(f"日期转换错误 ({date_val}): {str(e)}")
                    return date_val

            df['published_date'] = df['published_date'].apply(convert_date)
            print("统一 published_date 格式完成")

        # 新增：根据 published_date 创建 date 列
        if 'published_date' in df.columns:
            df['date'] = df['published_date'].apply(
                lambda x: str(x).split('_')[0] if pd.notna(x) and '_' in str(x) else (str(x).split(' ')[0] if pd.notna(x) and ' ' in str(x) else (str(x) if pd.notna(x) else pd.NA))
            )
            # 再次尝试转换为标准日期格式 YYYY-MM-DD，以防原始数据只有日期
            df['date'] = pd.to_datetime(df['date'], errors='coerce').dt.strftime('%Y-%m-%d')
        else:
            print("警告: 缺少 published_date 列，无法创建 date 列")
            df['date'] = pd.NA


        # 填充缺失列
        for col in MASTER_COLUMNS:
            if col not in df.columns:
                print(f"添加缺失列: {col}")
                df[col] = pd.NA
                
        # 列排序过滤
        df = df.reindex(columns=MASTER_COLUMNS)
        
        # 数据类型转换
        numeric_columns = ['like', 'comment', 'share', 'collect', 'subscribers', 'video_views']
        for col in numeric_columns:
            if col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('int64')
                except Exception as e:
                    print(f"转换列 {col} 时出错: {str(e)}")
        
        # 注意：根据要求，不再转换playthrough_rate，保留原始数据
        # 只进行基础类型处理，确保是字符串或数值类型
        if 'playthrough_rate' in df.columns:
            try:
                # 确保playthrough_rate列非空
                df['playthrough_rate'] = df['playthrough_rate'].fillna(0)
            except Exception as e:
                print(f"处理playthrough_rate列时出错: {str(e)}")
        
        # 处理avg_play_duration列：对于国内数据，需要去除"秒"字
        if 'avg_play_duration' in df.columns:
            try:
                # 创建一个新列用于存储处理后的值
                processed_values = []
                
                for idx, value in enumerate(df['avg_play_duration']):
                    try:
                        if pd.isna(value):
                            processed_values.append(0)
                            continue
                            
                        # 将值转换为字符串
                        value_str = str(value)
                        
                        # 对于包含"秒"的值(国内数据)，提取数字部分
                        if '秒' in value_str:
                            # 使用正则表达式提取数字部分，包括小数点
                            match = re.search(r'(\d+\.?\d*)', value_str)
                            if match:
                                processed_values.append(float(match.group(1)))
                            else:
                                processed_values.append(0)
                        else:
                            # 尝试直接转换为数值
                            try:
                                processed_values.append(float(value_str))
                            except:
                                # 如果无法转换，尝试提取数字部分
                                match = re.search(r'(\d+\.?\d*)', value_str)
                                if match:
                                    processed_values.append(float(match.group(1)))
                                else:
                                    processed_values.append(0)
                    except Exception as e:
                        print(f"处理第{idx}行avg_play_duration时出错: {str(e)}")
                        processed_values.append(0)
                
                # 更新列值
                df['avg_play_duration'] = processed_values
                print("avg_play_duration列处理完成")
                
            except Exception as e:
                print(f"处理avg_play_duration列时出错: {str(e)}")
        
        # 处理post列中的换行符，确保数据一致性
        if 'post' in df.columns:
            df['post'] = df['post'].astype(str).apply(lambda x: x.replace('\n', ' ').replace('\r', ''))
            print("处理post列中的换行符")
        
        # 在处理完其他列之后，添加game_label列
        if 'post' in df.columns:
            print("正在生成game_label...")
            df['game_label'] = df['post'].apply(determine_game_label)
            print("game_label生成完成")
        
        # 确保game_label在MASTER_COLUMNS中
        if 'game_label' not in MASTER_COLUMNS:
            MASTER_COLUMNS.append('game_label')
        
        print(f"处理完成，最终数据行数: {len(df)}")
        return df
        
    except Exception as e:
        print(f"处理 {file_path} 失败: {str(e)}")
        print("错误详情:")
        traceback.print_exc()
        return None

# 主处理流程
def main():
    try:
        # 使用新的函数获取应用程序路径
        app_path = get_application_path()
        print(f"应用程序路径: {app_path}")
        
        # 使用应用程序路径来定义文件夹路径
        EXCEL_DIR = os.path.join(app_path, "files")
        output_dir = os.path.join(app_path, "processed_files")
        
        print(f"输入文件夹路径: {EXCEL_DIR}")
        print(f"处理后文件夹路径: {output_dir}")
        
        # 确保文件夹存在
        if not os.path.exists(EXCEL_DIR):
            os.makedirs(EXCEL_DIR)
            print(f"创建输入文件夹: {EXCEL_DIR}")
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"创建处理后文件夹: {output_dir}")
        
        all_data = []
        processed_files = 0
        failed_files = 0
        
        # 获取所有需要处理的文件（同时支持.xlsx和.csv）
        excel_files = [f for f in os.listdir(EXCEL_DIR) 
                    if (f.endswith(".xlsx") or f.endswith(".csv")) 
                    and not f.startswith("merged_")]
        
        print(f"找到{len(excel_files)}个文件需要处理: {excel_files}")
        
        # 先处理国内文件，再处理国外文件（可以根据需要调整顺序）
        for is_foreign_process in [False, True]:
            process_type = "国外" if is_foreign_process else "国内"
            
            for filename in excel_files:
                # 检查当前文件是否符合当前处理类型
                if is_foreign_file(filename) == is_foreign_process:
                    file_path = os.path.join(EXCEL_DIR, filename)
                    print(f"\n开始处理{process_type}文件: {filename}")
                    
                    processed_df = process_excel(file_path)
                    if processed_df is not None:
                        all_data.append(processed_df)
                        processed_files += 1
                        print(f"已成功处理 {filename}，移动到已处理目录")
                        try:
                            os.rename(file_path, os.path.join(output_dir, filename))
                        except Exception as e:
                            print(f"移动文件失败: {str(e)}")
                    else:
                        failed_files += 1
                        print(f"处理 {filename} 失败，跳过")
        
        print(f"\n所有文件处理完成，成功处理{processed_files}个文件，失败{failed_files}个文件")
        
        if all_data:
            print("开始合并数据...")
            final_df = pd.concat(all_data, ignore_index=True)
            print(f"合并前总行数: {sum(len(df) for df in all_data)}")
            
            # 检查重复的post_id
            duplicate_count = final_df['post_id'].duplicated().sum()
            if duplicate_count > 0:
                print(f"发现{duplicate_count}个重复的post_id，进行数据合并")
                
                # 使用正确的列名
                numeric_cols = ['video_views', 'like', 'comment', 'share', 'collect']
                
                # 先获取重复的post_id
                duplicated_posts = final_df[final_df['post_id'].duplicated(keep=False)]
                unique_posts = final_df[~final_df['post_id'].duplicated(keep='first')]
                
                # 对重复的post_id的数值列求和
                summed_values = duplicated_posts.groupby('post_id')[numeric_cols].sum()
                
                # 更新unique_posts中对应行的数值
                for post_id in summed_values.index:
                    for col in numeric_cols:
                        unique_posts.loc[unique_posts['post_id'] == post_id, col] = summed_values.loc[post_id, col]
                
                final_df = unique_posts
                print(f"合并后行数: {len(final_df)}")
            
            # 生成输出路径
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(
                app_path,
                f"merged_data_{timestamp}.xlsx"
            )
            
            # 保存结果
            final_df.to_excel(output_path, index=False)
            print(f"合并完成！输出文件：{output_path}")
        else:
            print("\n没有需要处理的有效文件")
            
    except Exception as e:
        print(f"主处理流程发生错误: {str(e)}")
        traceback.print_exc()
    
    # 添加暂停，让用户看到结果
    print("\n处理完成。按回车键退出...")
    input()

if __name__ == "__main__":
    main()
