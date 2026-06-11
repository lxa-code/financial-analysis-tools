# -*- coding: utf-8 -*-
"""
月度财务报告生成向导
引导完成月度报告的生成
"""

from pathlib import Path
import subprocess
import sys

WORKSPACE = Path(r"c:\Users\lecoo\CodeBuddy\20260425100224\城发投资工作专区")
TEMPLATE_DIR = WORKSPACE / "03_模板工具"

def step1():
    """步骤1：准备数据"""
    print()
    print("=" * 50)
    print("步骤1: 准备数据")
    print("=" * 50)
    print(f"请打开并填写Excel模板:")
    print(f"  {TEMPLATE_DIR / '城发投资月度财务数据填充模板.xlsx'}")
    print()
    input("填写完成后按回车键继续...")

def step2():
    """步骤2：运行填充脚本"""
    print()
    print("=" * 50)
    print("步骤2: 运行填充脚本")
    print("=" * 50)
    
    script = TEMPLATE_DIR / "自动填充月度财务说明.py"
    print(f"运行: {script}")
    print()
    
    # 运行脚本
    try:
        subprocess.run([sys.executable, str(script)], check=True)
        print("\n填充完成!")
    except Exception as e:
        print(f"错误: {e}")
    
    input("\n按回车键继续...")

def step3():
    """步骤3：检查结果"""
    print()
    print("=" * 50)
    print("步骤3: 检查结果")
    print("=" * 50)
    print("生成的报告文件:")
    
    for f in Path(".").glob("*数据填充版.docx"):
        print(f"  - {f.name}")
    
    print()
    print("请检查报告内容是否正确")

def run_wizard():
    """运行向导"""
    print()
    print("=" * 50)
    print("城发投月度财务报告生成向导")
    print("=" * 50)
    
    step1()
    step2()
    step3()
    
    print()
    print("=" * 50)
    print("报告生成完成!")
    print("=" * 50)

if __name__ == "__main__":
    run_wizard()
