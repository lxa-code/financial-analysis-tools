# -*- coding: utf-8 -*-
"""
腾讯乐享 API 工具 v2
基于官方 OpenAPI 文档 (lexiangla.com/wiki/api/)
支持：团队管理、知识库管理、知识节点CRUD、AI搜索/问答、文件上传、通讯录等
access_token 自动缓存管理，有效期2小时
"""

import json
import os
import time
import requests

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "lexiang_config.json")
CACHE_PATH = os.path.join(os.path.dirname(__file__), ".lexiang_token_cache.json")


def _load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_cache():
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_cache(data):
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


def get_access_token():
    """获取 access_token，自动缓存，有效期2小时"""
    cache = _load_cache()
    now = time.time()

    if cache.get("access_token") and cache.get("expires_at", 0) > now + 60:
        return cache["access_token"]

    config = _load_config()
    url = config["token_url"]
    payload = {
        "grant_type": "client_credentials",
        "app_key": config["app_key"],
        "app_secret": config["app_secret"],
    }

    resp = requests.post(url, json=payload)
    if resp.status_code != 200:
        raise Exception(f"获取access_token失败: {resp.status_code} - {resp.text}")

    data = resp.json()
    cache["access_token"] = data["access_token"]
    cache["expires_at"] = now + data.get("expires_in", 7200)
    _save_cache(cache)

    return data["access_token"]


class LexiangAPI:
    """腾讯乐享 API 客户端"""

    def __init__(self, staff_id=None):
        self.config = _load_config()
        self.base_url = self.config["api_base"]
        self.staff_id = staff_id or self.config.get("staff_id")

    def _request(self, method, path, params=None, json_data=None, need_staff=False):
        token = get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=utf-8",
        }
        if need_staff and self.staff_id:
            headers["x-staff-id"] = self.staff_id
        url = f"{self.base_url}{path}"
        resp = requests.request(method, url, headers=headers, params=params, json=json_data)

        if resp.status_code == 401:
            cache = _load_cache()
            cache.pop("access_token", None)
            _save_cache(cache)
            token = get_access_token()
            headers["Authorization"] = f"Bearer {token}"
            resp = requests.request(method, url, headers=headers, params=params, json=json_data)

        ct = resp.headers.get("content-type", "")
        if "application/json" in ct:
            return resp.json()
        # 即使content-type不是json，也尝试解析为JSON
        try:
            return resp.json()
        except (json.JSONDecodeError, ValueError):
            return resp.text

    # ===== 团队管理 =====
    def get_teams(self, limit=20, page_token=""):
        """获取团队列表"""
        return self._request("GET", "/cgi-bin/v1/kb/teams",
                             params={"limit": limit, "page_token": page_token})

    def get_team_detail(self, team_id):
        """获取团队详情"""
        return self._request("GET", f"/cgi-bin/v1/kb/teams/{team_id}")

    # ===== 知识库管理 =====
    def get_spaces(self, team_id=None, limit=20):
        """获取知识库列表"""
        params = {"limit": limit}
        if team_id:
            params["team_id"] = team_id
        return self._request("GET", "/cgi-bin/v1/kb/spaces", params=params)

    def get_space_detail(self, space_id):
        """获取知识库详情"""
        return self._request("GET", f"/cgi-bin/v1/kb/spaces/{space_id}")

    # ===== 知识节点 =====
    def get_entries(self, space_id, parent_id=None, limit=20):
        """获取知识节点列表"""
        params = {"space_id": space_id, "limit": limit}
        if parent_id:
            params["parent_id"] = parent_id
        return self._request("GET", "/cgi-bin/v1/kb/entries", params=params)

    def get_entry_detail(self, entry_id):
        """获取知识节点详情"""
        return self._request("GET", f"/cgi-bin/v1/kb/entries/{entry_id}")

    def get_entry_content(self, entry_id, content_type="html"):
        """获取在线文档内容 (仅 page 类型)"""
        return self._request("GET", f"/cgi-bin/v1/kb/entries/{entry_id}/content",
                             params={"content_type": content_type})

    # ===== AI 搜索 =====
    def ai_search(self, query, targets=None, top_n=5, with_score=False):
        """
        AI 智能搜索
        targets: [{"type": "space", "id": "xxx"}, ...]  可选，不传则全站搜索
        """
        body = {"query": query, "top_n": top_n}
        if targets:
            body["targets"] = targets
        if with_score:
            body["with_score"] = True
        return self._request("POST", "/cgi-bin/v1/ai/search",
                             json_data=body, need_staff=True)

    # ===== AI 问答 =====
    def ai_qa(self, query, targets=None, research=False):
        """AI 智能问答，research=True 启用专业研究模式"""
        body = {"query": query, "research": research}
        if targets:
            body["targets"] = targets
        return self._request("POST", "/cgi-bin/v1/ai/qa",
                             json_data=body, need_staff=True)

    # ===== AI FAQ =====
    def get_ai_faqs(self, limit=20, page=1):
        """获取AI助手常见问题列表"""
        return self._request("GET", "/cgi-bin/v1/ai-faqs",
                             params={"limit": limit, "page": page})

    # ===== 知识库创建 =====
    def create_space(self, team_id, name, description=""):
        """创建知识库"""
        body = {
            "name": name,
            "team_id": team_id,
            "visible_type": 2,
        }
        if description:
            body["description"] = description
        return self._request("POST", "/cgi-bin/v1/kb/spaces",
                             json_data=body, need_staff=True)

    # ===== 知识节点创建 =====
    def create_entry(self, space_id, name, entry_type, parent_id=None):
        """创建知识节点 (folder/page/file)"""
        body = {
            "data": {
                "type": "kb_entry",
                "attributes": {"name": name, "entry_type": entry_type},
                "relationships": {
                    "space": {"data": {"type": "kb_space", "id": space_id}}
                }
            }
        }
        if parent_id:
            body["data"]["relationships"]["parent_entry"] = {
                "data": {"type": "kb_entry", "id": parent_id}
            }
        return self._request("POST", "/cgi-bin/v1/kb/entries",
                             json_data=body, need_staff=True)

    # ===== 文件上传 =====
    def get_upload_params(self, name, media_type="file"):
        """获取文件上传凭证 (步骤1: 获取state+COS凭证)"""
        return self._request("POST", "/cgi-bin/v1/kb/files/upload-params",
                             json_data={"name": name, "media_type": media_type},
                             need_staff=True)

    def upload_file_to_cos(self, upload_params, local_file_path):
        """
        步骤2: 上传文件到腾讯云COS
        upload_params: get_upload_params() 的返回结果
        """
        obj = upload_params.get("object", {})
        url = f"https://{obj['Bucket']}.cos.{obj['Region']}.myqcloud.com/{obj['key']}"
        headers = {
            "Authorization": obj["Authorization"],
            "x-cos-security-token": obj["XCosSecurityToken"],
        }
        with open(local_file_path, "rb") as f:
            resp = requests.put(url, headers=headers, data=f)
        return resp.status_code in (200, 204)

    def create_file_entry(self, state, space_id, name, parent_id=None):
        """
        步骤3: 使用上传凭证创建文件知识节点
        state: get_upload_params() 返回的 state 字段
        """
        params = {"state": state, "space_id": space_id}
        body = {
            "data": {
                "type": "kb_entry",
                "attributes": {"name": name, "entry_type": "file"},
                "relationships": {
                    "space": {"data": {"type": "kb_space", "id": space_id}}
                }
            }
        }
        if parent_id:
            body["data"]["relationships"]["parent_entry"] = {
                "data": {"type": "kb_entry", "id": parent_id}
            }
        return self._request("POST", "/cgi-bin/v1/kb/entries",
                             params=params, json_data=body, need_staff=True)

    def upload_file(self, space_id, local_file_path, file_name=None, parent_id=None):
        """
        完整三步上传流程
        返回: (成功标志, 结果信息)
        """
        import os as _os
        name = file_name or _os.path.basename(local_file_path)
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        # 确定 media_type
        media_map = {"mp4": "video", "mov": "video", "avi": "video",
                     "mp3": "audio", "wav": "audio", "wma": "audio"}
        media_type = media_map.get(ext, "file")

        # 步骤1: 获取上传凭证
        params_result = self.get_upload_params(name, media_type)
        if params_result.get("code") and params_result["code"] != 0:
            return False, f"获取上传凭证失败: {params_result.get('message')}"

        state = params_result.get("state")
        if not state:
            return False, f"未获取到state: {json.dumps(params_result, ensure_ascii=False)[:200]}"

        # 步骤2: 上传到COS
        ok = self.upload_file_to_cos(params_result, local_file_path)
        if not ok:
            return False, "上传到COS失败"

        # 步骤3: 创建知识节点
        entry_result = self.create_file_entry(state, space_id, name, parent_id)
        if entry_result.get("code") and entry_result["code"] != 0:
            return False, f"创建知识节点失败: {entry_result.get('message')}"

        entry_id = entry_result.get("data", {}).get("id", "")
        return True, entry_id


def quick_test():
    """快速测试连接"""
    try:
        token = get_access_token()
        print(f"[OK] access_token 获取成功")
        print(f"   Token: {token[:30]}...")
        return True
    except Exception as e:
        print(f"[FAIL] 连接失败: {e}")
        return False


if __name__ == "__main__":
    quick_test()
