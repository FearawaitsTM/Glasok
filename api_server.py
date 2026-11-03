"""Glasok HTTP service for this computer or a trusted local network."""
from __future__ import annotations

import argparse
import base64
import ctypes
import json
import os
import hashlib
import secrets
import sys
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

from api_sources import REPORTS, fetch_selected

HOST="127.0.0.1";PORT=int(os.environ.get('GLAZOK_PREVIEW_PORT','8765'))


def app_dir():
    root=Path(os.environ.get("APPDATA") or Path.home()/".config")/"Glasok"
    root.mkdir(parents=True,exist_ok=True);return root


def _protect(raw):
    if sys.platform!="win32":return b"plain:"+raw
    class BLOB(ctypes.Structure):_fields_=[("cbData",ctypes.c_ulong),("pbData",ctypes.POINTER(ctypes.c_char))]
    src=ctypes.create_string_buffer(raw);src_blob=BLOB(len(raw),ctypes.cast(src,ctypes.POINTER(ctypes.c_char)));dst=BLOB()
    encrypt=ctypes.windll.crypt32.CryptProtectData
    encrypt.argtypes=[ctypes.POINTER(BLOB),ctypes.c_wchar_p,ctypes.POINTER(BLOB),ctypes.c_void_p,ctypes.c_void_p,ctypes.c_ulong,ctypes.POINTER(BLOB)]
    encrypt.restype=ctypes.c_int
    free=ctypes.windll.kernel32.LocalFree;free.argtypes=[ctypes.c_void_p];free.restype=ctypes.c_void_p
    if not encrypt(ctypes.byref(src_blob),"Glasok",None,None,None,0,ctypes.byref(dst)):raise ctypes.WinError()
    try:return b"dpapi:"+ctypes.string_at(dst.pbData,dst.cbData)
    finally:free(dst.pbData)


def _unprotect(raw):
    prefix,payload=raw.split(b":",1)
    if prefix==b"plain":return payload
    if sys.platform!="win32":raise ValueError("Windows DPAPI data")
    class BLOB(ctypes.Structure):_fields_=[("cbData",ctypes.c_ulong),("pbData",ctypes.POINTER(ctypes.c_char))]
    src=ctypes.create_string_buffer(payload);src_blob=BLOB(len(payload),ctypes.cast(src,ctypes.POINTER(ctypes.c_char)));dst=BLOB()
    decrypt=ctypes.windll.crypt32.CryptUnprotectData
    decrypt.argtypes=[ctypes.POINTER(BLOB),ctypes.POINTER(ctypes.c_wchar_p),ctypes.POINTER(BLOB),ctypes.c_void_p,ctypes.c_void_p,ctypes.c_ulong,ctypes.POINTER(BLOB)]
    decrypt.restype=ctypes.c_int
    free=ctypes.windll.kernel32.LocalFree;free.argtypes=[ctypes.c_void_p];free.restype=ctypes.c_void_p
    if not decrypt(ctypes.byref(src_blob),None,None,None,None,0,ctypes.byref(dst)):raise ctypes.WinError()
    try:return ctypes.string_at(dst.pbData,dst.cbData)
    finally:free(dst.pbData)


def load_settings():
    path=app_dir()/"api.dat"
    if not path.exists():return {}
    try:return json.loads(_unprotect(base64.b64decode(path.read_bytes())).decode("utf-8"))
    except Exception:return {}


def save_settings(data):
    current=load_settings();allowed={}
    for key in ("ozon_client_id","ozon_api_key","wb_api_key"):
        value=str(data.get(key,"")).strip()
        allowed[key]=value or str(current.get(key,""))
    path=app_dir()/"api.dat";path.write_bytes(base64.b64encode(_protect(json.dumps(allowed).encode("utf-8"))))
    try:os.chmod(path,0o600)
    except OSError:pass


def connection_token():
    """A separate bearer token for users; marketplace keys never leave the server."""
    path=app_dir()/"connection.dat"
    if path.exists():
        try:return _unprotect(base64.b64decode(path.read_bytes())).decode('ascii')
        except Exception:pass
    token=secrets.token_urlsafe(32)
    path.write_bytes(base64.b64encode(_protect(token.encode('ascii'))))
    try:os.chmod(path,0o600)
    except OSError:pass
    return token


def update_file():return app_dir()/"updates"/"Глазок.exe"


def publish_update(source):
    """Admin GUI publishes a Windows executable, then switches it atomically."""
    source=Path(source)
    if source.suffix.lower()!='.exe' or not source.is_file():
        raise ValueError('Выберите собранный файл Глазок.exe')
    target=update_file();target.parent.mkdir(parents=True,exist_ok=True)
    temporary=target.with_suffix('.upload')
    digest=hashlib.sha256();size=0
    try:
        with source.open('rb') as src,temporary.open('wb') as dst:
            while chunk:=src.read(1024*1024):
                dst.write(chunk);digest.update(chunk);size+=len(chunk)
        if size<1024 or temporary.open('rb').read(2)!=b'MZ':
            raise ValueError('Файл не похож на Windows EXE')
        os.replace(temporary,target)
    finally:
        temporary.unlink(missing_ok=True)
    metadata={'sha256':digest.hexdigest(),'size':size,'filename':'Глазок.exe'}
    (target.parent/'manifest.json').write_text(json.dumps(metadata),encoding='utf-8')
    return metadata


def published_update():
    target=update_file();manifest=target.parent/'manifest.json'
    if not target.is_file() or not manifest.is_file():return None
    try:
        result=json.loads(manifest.read_text(encoding='utf-8'))
        if target.stat().st_size==result['size']:return result
    except (OSError,KeyError,ValueError):pass
    return None


def public_status(settings):
    return {"ozon":bool(settings.get("ozon_client_id") and settings.get("ozon_api_key")),"wb":bool(settings.get("wb_api_key"))}


class Handler(BaseHTTPRequestHandler):
    server_version="GlasokServer/1.0"
    def log_message(self,fmt,*args):
        message=f'{self.client_address[0]} — {fmt % args}'
        callback=getattr(self.server,'on_log',None)
        if callback:callback(message)
        elif getattr(self.server,"verbose",False):super().log_message(fmt,*args)
    def _authorized(self):
        # HTTP is intended for a trusted LAN, never an internet-facing server.
        supplied=self.headers.get('Authorization','')
        expected='Bearer '+self.server.connection_token
        if secrets.compare_digest(supplied,expected):return True
        self._send(401,{'error':'Укажите код подключения к серверу'})
        return False
    def _send(self,status,payload):
        raw=json.dumps(payload,ensure_ascii=False,default=str).encode("utf-8")
        self.send_response(status);self.send_header("Content-Type","application/json; charset=utf-8");self.send_header("Content-Length",str(len(raw)));self.end_headers();self.wfile.write(raw)
    def _body(self):
        size=int(self.headers.get("Content-Length","0") or 0)
        if size>5_000_000:raise ValueError("Запрос слишком большой")
        return json.loads(self.rfile.read(size).decode("utf-8")) if size else {}
    def do_GET(self):
        if urlsplit(self.path).path=="/health":return self._send(200,{"ok":True,"service":"Глазок Сервер"})
        if not self._authorized():return
        if self.path=="/catalog":return self._send(200,{"reports":REPORTS,"settings":public_status(load_settings())})
        if self.path=="/settings":return self._send(200,public_status(load_settings()))
        if self.path=="/updates/latest":return self._send(200,{"update":published_update()})
        if self.path=="/updates/download":
            metadata=published_update()
            if not metadata:return self._send(404,{'error':'Новая версия не загружена'})
            path=update_file()
            self.send_response(200);self.send_header('Content-Type','application/octet-stream')
            self.send_header('Content-Length',str(metadata['size']))
            self.end_headers()
            with path.open('rb') as source:
                while chunk:=source.read(1024*1024):self.wfile.write(chunk)
            return
        return self._send(404,{"error":"Не найдено"})
    def do_POST(self):
        if not self._authorized():return
        try:body=self._body()
        except Exception as exc:return self._send(400,{"error":str(exc)})
        if self.path=="/fetch":
            market=body.get("marketplace");reports=body.get("reports") or [];start=body.get("date_from");end=body.get("date_to")
            if market not in REPORTS:return self._send(400,{"error":"marketplace должен быть Ozon или WB"})
            if not reports:return self._send(400,{"error":"Не выбран ни один отчёт"})
            try:
                start_date=date.fromisoformat(str(start));end_date=date.fromisoformat(str(end))
                if start_date>end_date:raise ValueError
            except ValueError:return self._send(400,{"error":"Период должен быть в формате ГГГГ-ММ-ДД"})
            try:
                views,errors=fetch_selected(market,reports,load_settings(),start,end)
                if self.server.on_log:
                    self.server.on_log(f'{market}: отчётов {len(reports)}, таблиц {len(views)}, ошибок {len(errors)}')
                    for error in errors:self.server.on_log(str(error)[:300])
                return self._send(200,{"ok":not errors,"marketplace":market,"views":views,"errors":errors})
            except Exception as exc:
                if self.server.on_log:self.server.on_log(f'{market}: ошибка получения отчётов: {str(exc)[:300]}')
                return self._send(500,{"error":str(exc)})
        return self._send(404,{"error":"Не найдено"})


def create_server(host=HOST,port=PORT,on_log=None):
    server=ThreadingHTTPServer((host,port),Handler)
    server.daemon_threads=True;server.connection_token=connection_token();server.on_log=on_log
    return server


def main():
    parser=argparse.ArgumentParser();parser.add_argument("--port",type=int,default=PORT);parser.add_argument("--host",default=HOST);parser.add_argument("--verbose",action="store_true");args=parser.parse_args()
    try:server=create_server(args.host,args.port)
    except OSError as exc:
        print(f"Глазок Сервер не запущен: порт {args.port} недоступен. {exc}",flush=True)
        print("Если Глазок уже открыт, второй сервер запускать не нужно.",flush=True)
        raise SystemExit(2)
    server.verbose=args.verbose
    print(f"Глазок Сервер: http://{args.host}:{args.port}",flush=True)
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()


if __name__=="__main__":main()
