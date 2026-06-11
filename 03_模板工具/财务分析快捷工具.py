# -*- coding: utf-8 -*-
"""
城发投财务分析快捷工具
常用财务分析功能
"""

import pandas as pd
from pathlib import Path

# 城发投知识库路径
WORKSPACE = Path(r"c:\Users\lecoo\CodeBuddy\20260425100224\城发投资工作专区")
KNOWLEDGE_DIR = WORKSPACE / "04_知识库"

def load_data():
    """加载数据"""
    print("加载城发投数据...")
    
    # 知识库索引
    index_file = KNOWLEDGE_DIR / "文件索引.json"
    if index_file.exists():
        import json
        with open(index_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    print("请先运行数据索引工具")
    return None

def quick_analysis():
    """快速财务分析菜单"""
    print()
    print("=" * 40)
    print("城发投财务分析工具")
    print("=" * 40)
    print("1. 财务比率分析")
    print("2. 盈利能力分析")
    print("3. 偿债能力分析")
    print("4. 营运能力分析")
    print("5. 成长能力分析")
    print("6. 杜邦分析")
    print("0. 退出")
    print("=" * 40)
    
    choice = input("请选择: ")
    
    if choice == "1":
        print("财务比率分析功能开发中...")
    elif choice == "6":
        print("杜邦分析: ROE = 销售净利率 x 总资产周转率 x 权益乘数")
    elif choice == "0":
        return
    
    input("\n按回车键继续...")

if __name__ == "__main__":
    while True:
        quick_analysis()
