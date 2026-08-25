# -*- coding: utf-8 -*-
"""衣橱本地服务：静态文件 + POST /api/clothes 录入接口
用法：python serve.py  然后访问 http://localhost:8765
"""
import json
import os
import sys
import threading
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WARDROBE_FILE = os.path.join(BASE_DIR, "wardrobe.json")
HOST = "0.0.0.0"
PORT = 8765
LOCK = threading.Lock()


def load_wardrobe():
    with open(WARDROBE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_wardrobe(data):
    tmp = WARDROBE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, WARDROBE_FILE)


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def log_message(self, fmt, *args):
        sys.stdout.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), fmt % args))

    def do_GET(self):
        # 根路径返回 index.html
        if self.path in ("/", ""):
            self.path = "/index.html"
        return super().do_GET()

    def do_POST(self):
        if self.path != "/api/clothes":
            self.send_error(404, "Not Found")
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception as e:
            self._json_response(400, {"success": False, "message": "请求体解析失败: %s" % e})
            return

        name = (body.get("name") or "").strip()
        category = (body.get("category") or "").strip()
        if not name or not category:
            self._json_response(400, {"success": False, "message": "名称和类别不能为空"})
            return

        seasons = body.get("seasons") or []
        if isinstance(seasons, str):
            seasons = [s.strip() for s in seasons.split(",") if s.strip()]
        seasons = [s for s in seasons if s in ("春", "夏", "秋", "冬")]

        with LOCK:
            try:
                data = load_wardrobe()
                new_id = max((int(i.get("id") or 0) for i in data), default=0) + 1
                item = {
                    "id": new_id,
                    "name": name,
                    "category": category,
                    "color": (body.get("color") or "").strip(),
                    "style": (body.get("style") or "").strip(),
                    "brand": (body.get("brand") or "").strip(),
                    "material": (body.get("material") or "").strip(),
                    "seasons": seasons if seasons else ["春", "秋"],
                    "occasion": (body.get("occasion") or "").strip(),
                    "purchase_date": (body.get("purchase_date") or "").strip(),
                    "note": (body.get("note") or "").strip(),
                    "created_at": time.strftime("%Y-%m-%d"),
                    "image": "",
                }
                data.append(item)
                save_wardrobe(data)
            except Exception as e:
                self._json_response(500, {"success": False, "message": "写入失败: %s" % e})
                return

        self._json_response(200, {"success": True, "message": "已保存", "item": item})

    def _json_response(self, code, obj):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


if __name__ == "__main__":
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print("=" * 50)
    print("  My Wardrobe Server  /  My Closet")
    print("  URL : http://localhost:%d" % PORT)
    print("  Data: wardrobe.json")
    print("  Stop: Ctrl+C  or  close this window")
    print("=" * 50)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("server stopped.")
