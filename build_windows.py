"""Запускается батником внутри отдельного виртуального окружения."""
import subprocess
import sys
import shutil
from pathlib import Path

root=Path(__file__).resolve().parent
icon=root/'1.ico'
if not icon.is_file():
    from PIL import Image
    picture=Image.open(root/'glazok.png').convert('RGBA')
    picture.save(icon,format='ICO',sizes=[(16,16),(32,32),(48,48),(256,256)])
    print('Создана иконка 1.ico из glazok.png.',flush=True)
if icon.read_bytes()[:4]!=b'\x00\x00\x01\x00':
    raise SystemExit('1.ico должен быть настоящим ICO-файлом, а не переименованной PNG/JPG картинкой.')
common=[sys.executable,'-m','PyInstaller','--noconfirm','--clean','--onefile','--icon',str(icon)]
assets=['--add-data',str(icon)+';.','--add-data',str(root/'glazok.png')+';.']
subprocess.run(common+['--windowed','--name','Глазок Сервер']+assets+[str(root/'server_app.py')],cwd=root,check=True)
subprocess.run(common+['--windowed','--name','Глазок']+assets+['--add-data',str(root/'README.txt')+';.','--hidden-import','xlrd',str(root/'glasok.py')],cwd=root,check=True)
client=root/'dist'/'Глазок.exe';server=root/'dist'/'Глазок Сервер.exe'
if not client.is_file() or not server.is_file():raise SystemExit('Один из EXE не найден после сборки.')
# BAT remains ASCII and is written with Windows CRLF line endings. cmd.exe can
# otherwise execute chopped fragments of an LF-only batch as separate commands.
bat_text=(root/'START_GLAZOK.bat').read_text(encoding='ascii').replace('\r\n','\n').replace('\n','\r\n')
(root/'dist'/'START_GLAZOK.bat').write_bytes(bat_text.encode('ascii'))
shutil.copy2(root/'start_glazok.ps1',root/'dist'/'start_glazok.ps1')
shutil.copy2(root/'apply_update.ps1',root/'dist'/'apply_update.ps1')
shutil.copy2(root/'README.txt',root/'dist'/'README.txt')
print(f'Готово: {client}\nГотово: {server}',flush=True)
subprocess.Popen(['cmd','/c',str(root/'dist'/'START_GLAZOK.bat')],cwd=server.parent)
