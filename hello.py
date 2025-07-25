print("Hel
      lo Hubei University of Science")
#!/usr/bin/env python3
"""
湖北科技学院计算机专业 - 增强版Hello程序
支持多语言问候、学号验证和日志记录
"""

import argparse
import datetime
import logging

# 配置日志系统
logging.basicConfig(
    filename='hello.log',
    level=logging.INFO,
    format='%(asctime)s - %(message)s'
)

def validate_student_id(student_id: str) -> bool:
    """验证湖北科技学院学号格式（2025级示例）"""
    if len(student_id) != 8:
        return False
    if not student_id.startswith('2025'):
        return False
    return student_id[4:].isdigit()

def main():
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(
        description='湖北科技学院计算机专业问候程序',
        epilog='示例: python hello.py --name 张三 --id 20251234'
    )
    parser.add_argument('-n', '--name', required=True, help='你的姓名')
    parser.add_argument('-id', '--student_id', help='学号（格式: 2025XXXX）')
    parser.add_argument('-en', '--english', action='store_true', help='英文模式')

    args = parser.parse_args()

    # 验证学号
    if args.student_id and not validate_student_id(args.student_id):
        print(f"错误：学号 {args.student_id} 格式无效！")
        logging.error(f"无效学号: {args.student_id}")
        return

    # 生成问候语
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    greeting = ""

    if args.english:
        greeting = f"Hello {args.name}, welcome to Computer Science at Hubei University of Science!"
    else:
        greeting = f"你好{args.name}，欢迎来到湖北科技学院计算机科学与技术专业！"

    if args.student_id:
        greeting += f"\n学号: {args.student_id}"

    greeting += f"\n当前时间: {current_time}"

    print("\n" + "="*50)
    print(greeting)
    print("="*50)

    # 记录日志
    logging.info(f"用户 {args.name} (ID: {args.student_id or 'N/A'}) 执行程序")

if __name__ == "__main__":
    main()
