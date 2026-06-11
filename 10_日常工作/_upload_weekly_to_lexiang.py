# -*- coding: utf-8 -*-
"""
批量上传周报/周例会文档到腾讯乐享知识库

目标团队: 要素组 (team_id: 771632a45a4411f18c1ade1e476b8fe0)
知识库名: 周例会与工作周报
"""
import json, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from _lexiang_api import LexiangAPI

# 项目根目录
BASE = r"d:\BaiduSyncdisk\CODE\城发投资工作专区"

# 目标团队ID (要素组)
TEAM_ID = "771632a45a4411f18c1ade1e476b8fe0"

# 要上传的文件列表 (路径, 乐享中文件夹名)
FILES_TO_UPLOAD = [
    # 周报/日报数据源
    ("_weekly_content.txt", "数据源", "周报内容数据.txt"),
    ("_daily_content.txt", "数据源", "日报内容数据.txt"),
    # 领导要求
    ("10_日常工作/周工作要求/领导要求总结_20260520.md", "周工作要求", "领导要求总结_20260520.md"),
    # 周报模板
    ("10_日常工作/个人工作周报/个人周报模板.docx", "周报模板", "个人周报模板.docx"),
    ("10_日常工作/个人工作周报/2-1 城发投资周例会材料要素组模板.xlsx", "周报模板", "周例会材料要素组模板.xlsx"),
    # 已生成周报
    ("10_日常工作/个人工作周报/个人周报_20260525-0527.docx", "个人周报", "个人周报_20260525-0527.docx"),
    ("10_日常工作/2026 06 08 业务赋能专班周报.xlsx", "部门周报", "业务赋能专班周报_20260608.xlsx"),
    # 部门周报
    ("10_日常工作/公司各部门周报/六月第一周（第二十三周）.xlsx", "部门周报", "六月第一周_第二十三周.xlsx"),
]


def main():
    api = LexiangAPI()
    print("=" * 70)
    print("批量上传周报/周例会文档到腾讯乐享")
    print("=" * 70)

    # Step 1: 创建知识库 (如果已存在则跳过)
    print("\n[1/4] 检查/创建知识库...")
    spaces = api.get_spaces(team_id=TEAM_ID)
    if isinstance(spaces, str):
        spaces = json.loads(spaces)
    existing = spaces.get("data", [])
    space_id = None
    space_name = "周例会与工作周报"

    for s in existing:
        if s["attributes"]["name"] == space_name:
            space_id = s["id"]
            print(f"  知识库已存在: {space_name} (id: {space_id})")
            break

    if not space_id:
        print(f"  创建知识库: {space_name}...")
        result = api.create_space(TEAM_ID, space_name, "城发投资要素组周例会材料、工作周报归档")
        if isinstance(result, str):
            result = json.loads(result)
        if result.get("code") and result["code"] != 0:
            print(f"  [FAIL] 创建失败: {result.get('message')}")
            print(f"  完整响应: {json.dumps(result, ensure_ascii=False)[:500]}")
            return
        space_id = result.get("data", {}).get("id")
        print(f"  创建成功! space_id: {space_id}")
    else:
        space_id = space_id  # already set

    # Step 2: 创建文件夹结构
    print("\n[2/4] 创建文件夹结构...")
    folders_needed = sorted(set(f[1] for f in FILES_TO_UPLOAD))
    folder_ids = {}
    for folder_name in folders_needed:
        result = api.create_entry(space_id, folder_name, "folder")
        if isinstance(result, str):
            result = json.loads(result)
        fid = result.get("data", {}).get("id", "")
        if fid:
            folder_ids[folder_name] = fid
            print(f"  [OK] {folder_name} (id: {fid})")
        else:
            print(f"  [WARN] {folder_name} 创建可能失败: {result.get('message', result)[:100]}")

    # Step 3: 上传文件
    print("\n[3/4] 上传文件...")
    success_count = 0
    fail_count = 0

    for file_rel, folder, display_name in FILES_TO_UPLOAD:
        local_path = os.path.join(BASE, file_rel)
        if not os.path.exists(local_path):
            print(f"  [SKIP] 文件不存在: {file_rel}")
            fail_count += 1
            continue

        parent_id = folder_ids.get(folder)
        print(f"  上传: {display_name} -> {folder}...", end=" ")
        ok, result = api.upload_file(space_id, local_path, display_name, parent_id)
        if ok:
            print(f"[OK] entry_id: {result}")
            success_count += 1
        else:
            print(f"[FAIL] {result[:150]}")
            fail_count += 1

    # Step 4: 总结
    print(f"\n[4/4] 完成!")
    print(f"  成功: {success_count}  失败: {fail_count}")
    print(f"  知识库: {space_name} (id: {space_id})")
    print(f"  团队: 要素组 (id: {TEAM_ID})")


if __name__ == "__main__":
    main()
