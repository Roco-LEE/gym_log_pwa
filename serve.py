# -*- coding: utf-8 -*-
"""헬스앱 폴더를 같은 Wi-Fi 안의 폰에 보여주는 서버 + 엑셀 동기화 API.
   실행: python serve.py          →  폰 크롬에서 http://<이 PC IP>:8123
         python serve.py --test   →  동기화를 헬스일지_테스트.xlsx 에 (진짜 파일 안 건드림)
   POST /api/sync  {sessions:[...]}  →  헬스일지.xlsx [기록] 에 추가, {written:[id], rows:n} 응답"""
import http.server, socket, os, json, traceback, sys
import sync_excel
PORT = 8123
TEST = "--test" in sys.argv
os.chdir(os.path.dirname(os.path.abspath(__file__)))

class H(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map,
                      ".webmanifest": "application/manifest+json", ".js": "text/javascript", ".svg": "image/svg+xml"}
    def end_headers(self):
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()

    def do_POST(self):
        if self.path != "/api/sync":
            return self.send_error(404)
        try:
            n = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(n).decode("utf-8"))
            r = sync_excel.append_sessions(body.get("sessions", []), body.get("exMemo"))
            r["target"] = os.path.basename(sync_excel.XLSX)
            r["test"] = TEST   # 테스트면 앱이 "보냈음" 표시를 남기지 않음 (나중에 진짜 파일에 다시 보낼 수 있게)
            code, out = 200, r
            print(f"동기화 → {r['target']}: {r['rows']}줄 추가 (세션 {len(r['written'])}개), 운동 메모 {r.get('memos', 0)}개")
        except PermissionError:
            code, out = 409, {"error": "엑셀 파일이 열려 있어요. 닫고 다시 눌러주세요."}
        except Exception as e:
            traceback.print_exc()
            code, out = 500, {"error": str(e)}
        data = json.dumps(out, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

def my_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM); s.connect(("8.8.8.8", 80)); ip = s.getsockname()[0]; s.close(); return ip
    except Exception:
        return "127.0.0.1"

# 크롬이 미리 여는 빈 연결 때문에 단일 스레드 서버는 멈추므로 멀티스레드로
with http.server.ThreadingHTTPServer(("0.0.0.0", PORT), H) as httpd:
    if TEST:
        print("★ 테스트 모드: 동기화가", os.path.basename(sync_excel.use_test_file()), "에 들어감 (진짜 헬스일지.xlsx 는 안 건드림)")
    print(f"폰에서 열기:  http://{my_ip()}:{PORT}   (Ctrl+C 로 종료)")
    try: httpd.serve_forever()
    except KeyboardInterrupt: pass
