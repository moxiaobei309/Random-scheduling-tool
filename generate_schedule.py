import random
import pandas as pd
from datetime import datetime, timedelta
import argparse
import sys

def load_names(filename='names.txt'):
    """从文件加载名字列表"""
    try:
        # 使用绝对路径确保文件位置正确
        import os
        filepath = os.path.join(os.path.dirname(__file__), filename)
        with open(filepath, 'r', encoding='utf-8') as f:
            names = [line.strip() for line in f if line.strip()]
            if len(names) < 2:
                raise ValueError("至少需要2个名字")
            return names
    except FileNotFoundError:
        print(f"错误：找不到名字文件 {filename}")
        sys.exit(1)

def generate_dates(start_date, days):
    """生成日期范围"""
    return [start_date + timedelta(days=i) for i in range(days)]

def validate_input(names, days):
    """验证输入参数"""
    if days < 1:
        raise ValueError("天数必须大于0")
    if len(names) < 2:
        raise ValueError("至少需要2个名字")
    if days * 2 > len(names):
        print("警告：值班人数不足，将重复使用人员")

def generate_schedule(names, start_date, days):
    """生成值班表"""
    dates = generate_dates(start_date, days)
    schedule = pd.DataFrame(columns=['日期', '星期', '上午', '下午'])
    
    # 初始化值班统计
    duty_count = {name: 0 for name in names}
    last_duty = {name: -1 for name in names}  # 使用-1作为初始值
    
    # 打乱初始顺序
    shuffled_names = names.copy()
    random.shuffle(shuffled_names)
    
    for i, date in enumerate(dates):
        # 选择最少值班的人员
        candidates = [name for name in shuffled_names 
                     if last_duty.get(name, -7) < i - 1]  # 确保不连续值班
        
        # 按值班次数排序，选择最少值班的2人
        if len(candidates) < 2:
            # 如果没有足够候选人，重置last_duty
            candidates = names.copy()
            selected = sorted(candidates, key=lambda x: duty_count[x])[:2]
        else:
            selected = sorted(candidates, key=lambda x: duty_count[x])[:2]
        
        # 更新值班统计
        for name in selected:
            duty_count[name] += 1
            last_duty[name] = i
        
        # 打乱顺序以增加随机性
        random.shuffle(selected)
        
        new_row = pd.DataFrame({
            '日期': [date.strftime('%Y-%m-%d')],
            '星期': [['周一', '周二', '周三', '周四', '周五', '周六', '周日'][date.weekday()]],
            '上午': [selected[0]],
            '下午': [selected[1]]
        })
        schedule = pd.concat([schedule, new_row], ignore_index=True)
    
    # 添加值班统计
    stats = pd.DataFrame({
        '姓名': list(duty_count.keys()),
        '值班次数': list(duty_count.values())
    }).sort_values('值班次数', ascending=False)
    
    return schedule, stats

def main():
    parser = argparse.ArgumentParser(description='生成随机值班表')
    parser.add_argument('-s', '--start', type=str, default=None,
                       help='起始日期 (格式: YYYY-MM-DD)，默认为下周一开始')
    parser.add_argument('-d', '--days', type=int, default=5,
                       help='生成天数，默认为5天')
    parser.add_argument('-o', '--output', type=str, default='值班表.xlsx',
                       help='输出文件名，默认为 值班表.xlsx')
    parser.add_argument('-n', '--names', type=str, default='names.txt',
                       help='名字列表文件，默认为 names.txt')
    
    args = parser.parse_args()
    
    # 加载名字
    names = load_names(args.names)
    
    # 设置起始日期
    start_date = (datetime.strptime(args.start, '%Y-%m-%d') if args.start 
                 else datetime.now() + timedelta(days=(7 - datetime.now().weekday())))
    
    # 验证输入
    try:
        validate_input(names, args.days)
    except ValueError as e:
        print(f"输入错误：{e}")
        sys.exit(1)
    
    # 生成值班表
    schedule, stats = generate_schedule(names, start_date, args.days)
    
    # 保存到Excel
    with pd.ExcelWriter(args.output) as writer:
        schedule.to_excel(writer, sheet_name='值班表', index=False)
        stats.to_excel(writer, sheet_name='值班统计', index=False)
    print(f"成功生成值班表：{args.output}，包含值班统计信息")

if __name__ == '__main__':
    main()
    input("按回车键退出...")  # 添加暂停，防止窗口自动关闭
