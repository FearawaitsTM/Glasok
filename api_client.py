"""Connection to a local or LAN Glasok server."""
import hashlib
import base64
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from urllib.error import HTTPError,URLError
from urllib.parse import urlsplit
from urllib.request import Request,urlopen

BASE='http://127.0.0.1:'+os.environ.get('GLAZOK_PREVIEW_PORT','8765')


def client_config_path():
    directory=Path(os.environ.get('APPDATA') or Path.home()/'.config')/'Glasok'
    directory.mkdir(parents=True,exist_ok=True)
    return directory/'client_connection.json'


def connection():
    if os.environ.get('GLAZOK_PREVIEW_PORT'):
        from api_server import connection_token
        return BASE,connection_token()
    try:
        saved=json.loads(client_config_path().read_text(encoding='utf-8'))
        from api_server import _unprotect
        token=_unprotect(base64.b64decode(saved['token_secret'])).decode('utf-8')
        return saved['address'].rstrip('/'),token
    except (OSError,ValueError,KeyError):
        # On one computer the server and client share the local DPAPI secret.
        from api_server import connection_token
        return BASE,connection_token()


def save_connection(address,token):
    address=address.strip().rstrip('/');token=token.strip()
    parsed=urlsplit(address)
    if parsed.scheme not in ('http','https') or not parsed.hostname or parsed.path or parsed.query or parsed.fragment or not token:
        raise ValueError('Введите адрес вида http://192.168.1.10:8765 и код подключения')
    from api_server import _protect
    encrypted=base64.b64encode(_protect(token.encode('utf-8'))).decode('ascii')
    path=client_config_path();path.write_text(json.dumps({'address':address,'token_secret':encrypted}),encoding='utf-8')
    try:os.chmod(path,0o600)
    except OSError:pass


def _request(path,body=None,timeout=600):
    address,token=connection()
    data=None if body is None else json.dumps(body,ensure_ascii=False).encode('utf-8')
    return Request(address+path,data=data,method='POST' if data is not None else 'GET',
                   headers={'Content-Type':'application/json','Authorization':'Bearer '+token}),timeout


def request(path,body=None,timeout=600):
    req,timeout=_request(path,body,timeout)
    try:
        with urlopen(req,timeout=timeout) as response:return json.loads(response.read().decode('utf-8'))
    except HTTPError as exc:
        try:detail=json.loads(exc.read().decode('utf-8')).get('error')
        except Exception:detail=exc.reason
        raise RuntimeError(f'Сервер: {detail}') from exc
    except URLError as exc:raise RuntimeError('Глазок Сервер недоступен') from exc


def ensure_server():
    address,_=connection()
    try:request('/health',timeout=2);return True
    except Exception:pass
    if address!=BASE:return False
    root=Path(getattr(sys,'_MEIPASS',Path(__file__).resolve().parent))
    if getattr(sys,'frozen',False):
        exe=Path(sys.executable).resolve().parent/'Глазок Сервер.exe'
        if not exe.exists():return False
        subprocess.Popen([str(exe)],cwd=exe.parent)
    else:subprocess.Popen([sys.executable,str(root/'server_app.py')],cwd=root)
    for _ in range(30):
        time.sleep(.2)
        try:request('/health',timeout=2);return True
        except Exception:pass
    return False


def download_update(metadata):
    """Stage the published EXE only after verifying its advertised hash and size."""
    if sys.platform!='win32' or not getattr(sys,'frozen',False):return None
    expected=metadata.get('sha256','')
    if len(expected)!=64 or not all(c in '0123456789abcdef' for c in expected):raise ValueError('Некорректная контрольная сумма')
    executable=Path(sys.executable).resolve()
    digest=hashlib.sha256()
    with executable.open('rb') as current:
        while chunk:=current.read(1024*1024):digest.update(chunk)
    if digest.hexdigest()==expected:return None
    target=executable.with_suffix('.exe.new')
    if target.exists():
        old=hashlib.sha256(target.read_bytes()).hexdigest()
        if old==expected:return target
    request_obj,_=_request('/updates/download',timeout=180)
    temporary=target.with_suffix('.download');size=0;digest=hashlib.sha256()
    try:
        with urlopen(request_obj,timeout=180) as response,temporary.open('wb') as dst:
            while chunk:=response.read(1024*1024):
                dst.write(chunk);digest.update(chunk);size+=len(chunk)
        if size!=metadata['size'] or digest.hexdigest()!=expected:raise ValueError('Загруженный EXE не прошёл проверку')
        os.replace(temporary,target)
    finally:temporary.unlink(missing_ok=True)
    return target


def launch_update(staged):
    executable=Path(sys.executable).resolve()
    script=executable.parent/'apply_update.ps1'
    if not script.is_file():raise RuntimeError('Рядом с приложением нет apply_update.ps1')
    subprocess.Popen(['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',str(script),
                      '-Target',str(executable),'-Staged',str(staged),'-WaitPid',str(os.getpid())],
                     cwd=executable.parent,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
