# -*- coding: utf-8 -*-
import sys, json, os
sys.path.insert(0, os.path.dirname(__file__))
from _lexiang_api import LexiangAPI

api = LexiangAPI()

# 获取所有团队(翻页)
print("查找 team code k100197 和 k100022, k100032, k100059 对应的团队名")
print()

# 尝试通过code查询
for code in ["k100197", "k100022", "k100032", "k100059"]:
    r = api._request("GET", f"/cgi-bin/v1/teams?code={code}", need_staff=True)
    if isinstance(r, dict):
        data = r.get("data", [])
        for t in data:
            attrs = t.get("attributes", {})
            print(f"  [{attrs.get('code')}] {attrs.get('name')} (id={t.get('id')}) type={attrs.get('type')} secret={attrs.get('is_secret')}")
    else:
        print(f"  {code}: {type(r).__name__}")
