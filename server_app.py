"""Desktop administration window for the Glasok report and update server."""
import json
import ipaddress
import os
import queue
import socket
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog,messagebox,ttk
from urllib.request import urlopen

from api_server import PORT,app_dir,connection_token,create_server,publish_update,public_status,load_settings,save_settings,published_update
from app_icon import set_app_icon


def local_address():
    if sys.platform=='darwin':
        for interface in ('en0','en1'):
            try:
                result=subprocess.run(['ipconfig','getifaddr',interface],capture_output=True,text=True,timeout=2)
                value=result.stdout.strip()
                if result.returncode==0 and is_lan_address(value):return value
            except (OSError,subprocess.TimeoutExpired):pass
    sock=socket.socket(socket.AF_INET,socket.SOCK_DGRAM)
    try:
        sock.connect(('192.0.2.1',80))
        value=sock.getsockname()[0]
        if is_lan_address(value):return value
    except OSError:
        pass
    finally:sock.close()
    try:
        value=socket.gethostbyname(socket.gethostname())
        if is_lan_address(value):return value
    except OSError:pass
    return 'проверьте IP в настройках Wi-Fi'


def is_lan_address(value):
    try:address=ipaddress.ip_address(value)
    except ValueError:return False
    return (address.version==4 and address.is_private and not address.is_loopback
            and address not in ipaddress.ip_network('198.18.0.0/15'))


SERVER_HELP="""ГЛАЗОК СЕРВЕР — КАК ПОЛЬЗОВАТЬСЯ

1. Статус «Работает» означает, что приложение может обратиться к серверу. Отчёты из Ozon/WB появятся после сохранения действующих ключей с нужными правами.
2. Для Ozon заполните оба поля: Client-Id и Api-Key. Для WB достаточно токена WB. Можно настроить один маркетплейс или оба.
3. Вставьте ключ сочетанием ⌘V на Mac или Ctrl+V на Windows. Рядом с полем есть кнопка «Вставить». После ввода нажмите «Сохранить ключи». Строка ниже покажет, какой маркетплейс настроен.
4. В «Глазке» выберите Ozon или WB, отметьте отчёты, укажите даты и нажмите «Получить выбранное». Если прав или данных нет, причина появится в приложении и журнале сервера.
5. Галочка «Разрешить подключение из локальной сети» нужна лишь для других компьютеров. После изменения выключите и снова запустите сервер. На этом Mac используется адрес 127.0.0.1.
6. Адрес и код подключения нужны клиентам на других компьютерах. Код не является ключом Ozon/WB; передавайте его только своим пользователям.
7. Кнопка публикации EXE работает с собранной Windows-версией. Для просмотра приложения на Mac она не требуется.
8. «Порт занят» значит, что сервер уже открыт в другом процессе либо порт занят другой программой. «Глазок» может показывать «сервер доступен» из первого процесса. Новый журнал получает записи только от своего процесса.
9. Если все API-отчёты пишут про сертификат HTTPS, соединение с площадкой не прошло проверку безопасности. Проверьте системные сертификаты Mac и настройки VPN/прокси. При необходимости попросите администратора установить доверенный сертификат сети в macOS; проверку HTTPS отключать не нужно.
"""


class ServerWindow:
    def __init__(self,root):
        self.root=root;self.server=None;self.events=queue.Queue()
        root.title('Глазок Сервер');root.geometry('870x700');root.minsize(680,540)
        set_app_icon(root)
        root.configure(bg='#dff4ff')
        style=ttk.Style()
        if 'clam' in style.theme_names():style.theme_use('clam')
        style.configure('TFrame',background='#dff4ff');style.configure('TLabel',background='#dff4ff',foreground='#123149')
        style.configure('TButton',padding=7,background='#079bea',foreground='white')
        style.configure('TLabelframe',background='#dff4ff');style.configure('TLabelframe.Label',background='#dff4ff',foreground='#006da8')
        top=ttk.Frame(root,padding=12);top.pack(fill='x')
        self.status=ttk.Label(top,text='Сервер выключен');self.status.pack(side='left')
        ttk.Button(top,text='ⓘ Как пользоваться',command=self.show_help).pack(side='right',padx=7)
        self.start_button=ttk.Button(top,text='Запустить',command=self.start);self.start_button.pack(side='right')
        self.stop_button=ttk.Button(top,text='Остановить',command=self.stop,state='disabled');self.stop_button.pack(side='right',padx=6)
        frame=ttk.LabelFrame(root,text='Подключение пользователей',padding=10);frame.pack(fill='x',padx=12,pady=5)
        saved=app_dir()/'server_ui.json'
        try:lan=json.loads(saved.read_text(encoding='utf-8')).get('lan',False)
        except (OSError,ValueError):lan=False
        self.preview=bool(os.environ.get('GLAZOK_PREVIEW_PORT'))
        self.lan=tk.BooleanVar(value=lan and not self.preview)
        ttk.Checkbutton(frame,text='Разрешить подключение из локальной сети',variable=self.lan,state='disabled' if self.preview else 'normal').pack(anchor='w')
        address=(f'Просмотр на Mac · только на этом ПК: http://127.0.0.1:{PORT}' if self.preview
                 else f'Адрес для пользователей: {local_address()}:{PORT}  ·  на этом ПК: http://127.0.0.1:{PORT}')
        ttk.Label(frame,text=address).pack(anchor='w',pady=(6,2))
        row=ttk.Frame(frame);row.pack(fill='x')
        ttk.Label(row,text='Код подключения:').pack(side='left')
        self.token=tk.StringVar(value=connection_token())
        ttk.Entry(row,textvariable=self.token,state='readonly',width=58).pack(side='left',padx=8)
        ttk.Button(row,text='Скопировать',command=lambda:(root.clipboard_clear(),root.clipboard_append(self.token.get()))).pack(side='left')
        ttk.Label(frame,text='Передайте адрес и код только своим пользователям. Для доступа через интернет нужен защищённый HTTPS-шлюз.').pack(anchor='w',pady=(5,0))
        keys=ttk.LabelFrame(root,text='Ключи маркетплейсов (хранятся только на сервере)',padding=10);keys.pack(fill='x',padx=12,pady=5)
        self.inputs={}
        for code,label in [('ozon_client_id','Ozon Client-Id'),('ozon_api_key','Ozon Api-Key'),('wb_api_key','Wildberries API токен')]:
            row=ttk.Frame(keys);row.pack(fill='x',pady=3)
            ttk.Label(row,text=label,width=23).pack(side='left')
            var=tk.StringVar();self.inputs[code]=var
            entry=ttk.Entry(row,textvariable=var,show='•' if code!='ozon_client_id' else '')
            entry.pack(side='left',fill='x',expand=True)
            entry.bind('<Control-v>',lambda event,field=entry:self.paste_key(field))
            entry.bind('<Control-V>',lambda event,field=entry:self.paste_key(field))
            entry.bind('<Command-v>',lambda event,field=entry:self.paste_key(field))
            entry.bind('<Command-V>',lambda event,field=entry:self.paste_key(field))
            ttk.Button(row,text='Вставить',command=lambda field=entry:self.paste_key(field)).pack(side='left',padx=(5,0))
        ttk.Label(keys,text='На Mac вставка: ⌘V или кнопка «Вставить». Ozon: нужны Client-Id и Api-Key; WB: токен.',wraplength=800).pack(anchor='w',pady=(7,0))
        ttk.Button(keys,text='Сохранить ключи',command=self.save_keys).pack(anchor='e',pady=5)
        self.keys_state=ttk.Label(keys,text='');self.keys_state.pack(anchor='w')
        self.update_keys_state()
        update=ttk.LabelFrame(root,text='Обновление приложения у подключённых пользователей',padding=10);update.pack(fill='x',padx=12,pady=5)
        ttk.Button(update,text='Выбрать и опубликовать Глазок.exe',command=self.upload).pack(side='left')
        self.update_state=ttk.Label(update,text=self.update_label());self.update_state.pack(side='left',padx=12)
        ttk.Label(root,text='Журнал сервера',padding=(12,5)).pack(anchor='w')
        self.logs=tk.Text(root,height=12,bg='white',fg='#123149',state='disabled',wrap='word')
        self.logs.pack(fill='both',expand=True,padx=12,pady=(0,12))
        root.protocol('WM_DELETE_WINDOW',self.close)
        root.after(150,self.drain)
        self.start()

    def log(self,message):self.events.put(message)
    def paste_key(self,entry):
        try:value=self.root.clipboard_get().strip()
        except tk.TclError:
            self.root.bell();return 'break'
        try:entry.delete('sel.first','sel.last')
        except tk.TclError:pass
        entry.insert('insert',value)
        return 'break'
    def show_help(self):
        win=tk.Toplevel(self.root);win.title('Глазок Сервер — информация');win.geometry('760x570');win.transient(self.root)
        body=ttk.Frame(win,padding=12);body.pack(fill='both',expand=True)
        text=tk.Text(body,wrap='word',bg='white',fg='#123149',font=('Arial',11),padx=14,pady=12)
        scroll=ttk.Scrollbar(body,orient='vertical',command=text.yview)
        text.configure(yscrollcommand=scroll.set);text.pack(side='left',fill='both',expand=True);scroll.pack(side='right',fill='y')
        text.insert('1.0',SERVER_HELP);text.configure(state='disabled')
    def drain(self):
        while not self.events.empty():
            line=f'{datetime.now():%H:%M:%S}  {self.events.get_nowait()}\n'
            self.logs.configure(state='normal');self.logs.insert('end',line)
            self.logs.see('end');self.logs.configure(state='disabled')
        self.root.after(150,self.drain)
    def update_keys_state(self):
        status=public_status(load_settings())
        self.keys_state.configure(text='Ozon: '+('настроен' if status['ozon'] else 'нет ключей')+'    ·    WB: '+('настроен' if status['wb'] else 'нет токена'))
    def save_keys(self):
        data={key:var.get().strip() for key,var in self.inputs.items()}
        if not any(data.values()):return messagebox.showinfo('Ключи','Заполните хотя бы одно поле. Пустые поля сохраняют прежние значения.')
        try:save_settings(data)
        except Exception as exc:return messagebox.showerror('Ключи',str(exc))
        for var in self.inputs.values():var.set('')
        self.update_keys_state();self.log('Ключи сохранены')
    def update_label(self):
        item=published_update()
        return 'Версия: '+item['sha256'][:12] if item else 'EXE ещё не загружен'
    def upload(self):
        path=filedialog.askopenfilename(title='Новая версия Глазок.exe',filetypes=[('Приложение Windows','*.exe')])
        if not path:return
        try:
            item=publish_update(path)
            self.update_state.configure(text=self.update_label())
            self.log(f'Опубликовано обновление {item["sha256"][:12]} ({item["size"]} байт)')
            messagebox.showinfo('Готово','Новая версия доступна клиентам. Приложения загрузят её и предложат перезапуск.')
        except Exception as exc:messagebox.showerror('Обновление',str(exc))
    def start(self):
        if self.server:return
        host='0.0.0.0' if self.lan.get() and not self.preview else '127.0.0.1'
        try:self.server=create_server(host,PORT,self.log)
        except OSError as exc:
            try:
                with urlopen(f'http://127.0.0.1:{PORT}/health',timeout=1) as response:
                    other=json.load(response).get('service')=='Глазок Сервер'
            except Exception:other=False
            if other:
                self.status.configure(text='Сервер уже работает в другом окне')
                self.log('Этот экземпляр не запущен: сервер уже использует порт. Ключи здесь сохраняются, журнал — в первом окне.')
            else:
                self.log(f'Порт {PORT} недоступен: {exc}');self.status.configure(text='Не запущен: порт занят')
            return
        if not self.preview:
            (app_dir()/'server_ui.json').write_text(json.dumps({'lan':self.lan.get()}),encoding='utf-8')
        threading.Thread(target=self.server.serve_forever,daemon=True).start()
        self.start_button.configure(state='disabled');self.stop_button.configure(state='normal')
        self.status.configure(text=f'Работает на {host}:{PORT}');self.log(f'Сервер запущен на {host}:{PORT}')
    def stop(self):
        if not self.server:return
        self.server.shutdown();self.server.server_close();self.server=None
        self.start_button.configure(state='normal');self.stop_button.configure(state='disabled')
        self.status.configure(text='Сервер выключен');self.log('Сервер остановлен')
    def close(self):
        self.stop();self.root.destroy()


if __name__=='__main__':
    root=tk.Tk();ServerWindow(root);root.mainloop()
