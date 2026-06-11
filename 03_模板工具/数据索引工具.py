# -*- coding: utf-8 -*-
"""
城发投知识库索引工具
快速检索知识库中的文件
"""

import json
import os
from pathlib import Path

def build_index():
    """构建文件索引"""
    workspace = Path(r"c:\Users\lecoo\CodeBuddy\20260425100224\城发投资工作专区")
    index = {}
    
    for dir_name in workspace.iterdir():
        if dir_name.is_dir() and not dir_name.name.startswith('.'):
            files = []
            for f in dir_name.rglob("*"):
                if f.is_file():
                    files.append({
                        "name": f.name,
                        "path": str(f.relative_to(workspace)),
                        "size": f.stat().st_size,
                    })
            index[dir_name.name] = files
    
    # 保存索引
    index_file = workspace / "文件索引.json"
    with open(index_file, 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
    
    print(f"索引已保存: {index_file}")
    return index

if __name__ == "__main__":
    build_index()
