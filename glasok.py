"""Глазок 4: файлы, локальный API-сервер и Excel-подобные таблицы."""
import sys
import threading
import queue
from datetime import date,timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
import pandas as pd
from excel_grid import ExcelGrid,display
from table_view import present, product_matches, product_suggestions, exact_product, apply_column_filters
from unified_engine import analyze_all,SUPPORTED,totals
from ozon_engine import visible
from api_client import BASE,ensure_server,request as api_request,connection,save_connection,download_update,launch_update
from api_sources import REPORTS
from app_icon import set_app_icon


HELP_TEXT = """ГЛАЗОК — КНОПКИ И ФУНКЦИИ

ФАЙЛЫ
«Добавить файлы Ozon / WB» — выберите один или несколько XLSX, XLSM, XLS, CSV, TSV или ZIP. Приложение определяет маркетплейс по содержимому. Повторное добавление одинакового файла не удваивает суммы.
«Добавить папку» — загружает поддерживаемые файлы из выбранной папки и её подпапок.
«Очистить всё» — удаляет загруженные из текущего окна данные и результаты API. Файлы на диске не удаляются.
Вкладки Ozon и WB — переключают таблицы двух маркетплейсов. Если формат не распознан, появится вкладка «Не распознано» с исходными строками. Ошибки загрузки выводятся сообщением; их список также доступен в этом окне справки.

ТАБЛИЦЫ И ПОИСК
Список над таблицей — выбирает вид: общая сводка, периоды, остатки, исходные строки и другие доступные виды.
«Общая сводка товаров» объединяет показатели одного товара за один период. «Все источники вместе» показывает ручные файлы рядом с API-данными.
«Итоги по периодам» суммируют подходящие количества и суммы. «Пересечения дат» показывают дневные данные отдельно от месячных. «Остатки + маркировка + план поставки» показывает состояние склада и оценку пополнения. «Поисковые запросы», «Финансы магазина» и виды «Исходные» сохраняют детализацию своих отчётов.
Поле поиска — сразу после нажатия показывает первые пять товаров. Ввод артикула, SKU или названия сразу убирает неподходящие строки; из подсказки можно выбрать один товар.
«Найти» — применяет введённый запрос; при пустом поле показывает первые пять товаров.
«Сбросить» — очищает общий поиск и все фильтры столбцов.
«Скрыть нули» — скрывает нулевые/пустые показатели и неактивные строки. Снятие галочки показывает полные данные.
«Расчёты» — показывает или скрывает вычисленные поля. Они окрашены фиолетовым; исходные значения остаются голубыми.
«Сохранить» — выгружает видимые строки в XLSX или CSV с учётом поиска и фильтров.
Кнопка ▼ в заголовке — фильтр конкретного столбца. Внутри есть поиск, галочки, «Все», «Снять» и «Применить».
«Сортировка» и стрелки ↑/↓ — выбирают столбец и порядок сортировки.
«Авто» — подбирает ширину выбранного столбца по тексту. Повторное нажатие возвращает прежнюю ширину. Двойной щелчок по границе буквы столбца действует так же.
Границы в строке букв меняют ширину столбцов; границы у номеров строк меняют высоту. Выделение и Ctrl+C или ⌘C копируют диапазон ячеек. Строка над таблицей показывает полный текст выбранной ячейки.

ДАННЫЕ И РАСЧЁТЫ
«Продажи и текущие остатки» открывается первым при наличии двух отчётов Ozon.
Сначала идут артикул, SKU, название и показатели в штуках. «Остаток, шт» берётся из доступного к продаже остатка отчёта управления остатками, а «Итог после возвратов, шт» = передано в доставку − возвращено за период.
Продажи в день = доставлено / количество календарных дней периода включительно. Остаток в днях = остаток в штуках / продажи в день. При нулевых продажах запас в днях не определён: ячейка пустая.
Финансовый блок начинается с выручки от доставленных товаров. Средняя цена продажи = выручка / доставлено; прибыль на единицу = прибыль из отчёта / доставлено. Себестоимость и состав прибыли остаются такими, как в исходном отчёте.
Деньги, возвращённые покупателям, не равны расходам на обработку возвратов или обратную логистику. Если исходник не содержит суммы возврата покупателю, она не вычисляется по средней цене.
Артикул, SKU и название товара стоят первыми; затем идут значения из отчётов, далее поля «Расчёт». Товары совмещаются по идентификаторам и совпадающим периодам. Разные даты сохраняются отдельно.
Заказы/выкупы в день = число за период ÷ число дней. Доли и конверсии рассчитываются при наличии числителя и ненулевого знаменателя.
План поставки на 30 дней = среднесуточные продажи × 30 минус текущий остаток и, если известны, товары в пути. Потребность не бывает отрицательной. Отдельно показан план на 28 дней.
Общие финансы магазина не приписываются отдельным товарам; месячные итоги не распределяются искусственно по дням. Пустая ячейка означает отсутствие данных, ноль — известное значение.

"""


class ReportPanel(ttk.Frame):
    def __init__(self,parent,market=None,api_callback=None):
        super().__init__(parent);self.views={};self.filtered=pd.DataFrame();self.column_filters={};self.selected_product=None;self.suggestions=[];self.suggest_popup=None;self._choosing_suggestion=False;self._pending_search=None
        if market:
            api=ttk.LabelFrame(self,text=f'{market} API — отметьте любые отчёты',padding=7);api.pack(fill='x',pady=(4,7))
            self.api_vars={}
            for code,title in REPORTS[market].items():
                var=tk.BooleanVar(value=code in ('stocks','orders'));self.api_vars[code]=var
                ttk.Checkbutton(api,text=title,variable=var).pack(side='left',padx=(0,8))
            today=date.today();self.date_from=tk.StringVar(value=str(today-timedelta(days=30)));self.date_to=tk.StringVar(value=str(today))
            ttk.Label(api,text='с').pack(side='left');ttk.Entry(api,textvariable=self.date_from,width=11).pack(side='left',padx=3)
            ttk.Label(api,text='по').pack(side='left');ttk.Entry(api,textvariable=self.date_to,width=11).pack(side='left',padx=3)
            self.api_button=ttk.Button(api,text='Получить выбранное',command=lambda:api_callback(market,[k for k,v in self.api_vars.items() if v.get()],self.date_from.get(),self.date_to.get()))
            self.api_button.pack(side='right')
        bar=ttk.Frame(self);bar.pack(fill='x',pady=(4,8))
        self.choice=ttk.Combobox(bar,state='readonly',width=42);self.choice.pack(side='left',fill='x',expand=True);self.choice.bind('<<ComboboxSelected>>',self.change_view)
        self.query=tk.StringVar();self.search_entry=ttk.Entry(bar,textvariable=self.query,width=25);self.search_entry.pack(side='left',padx=6)
        self.query.trace_add('write',self.on_query_change)
        self.search_entry.bind('<Return>',self.search_submit)
        self.search_entry.bind('<Escape>',lambda e:self.dismiss_suggestions())
        self.search_entry.bind('<FocusIn>',lambda e:self.after_idle(self.show_suggestions))
        self.search_entry.bind('<ButtonRelease-1>',lambda e:self.after_idle(self.show_suggestions))
        ttk.Button(bar,text='Найти',command=self.search_submit).pack(side='left');ttk.Button(bar,text='Сбросить',command=self.reset_search).pack(side='left',padx=3)
        self.hide=tk.BooleanVar(value=True);ttk.Checkbutton(bar,text='Скрыть нули',variable=self.hide,command=self.refresh).pack(side='left',padx=5)
        self.show_calculations=tk.BooleanVar(value=True);ttk.Checkbutton(bar,text='Расчёты',variable=self.show_calculations,command=self.refresh).pack(side='left',padx=5)
        ttk.Button(bar,text='Сохранить',command=self.export).pack(side='right')
        second=ttk.Frame(self);second.pack(fill='x',pady=(0,6))
        ttk.Label(second,text='Сортировка:').pack(side='left');self.sortcol=ttk.Combobox(second,state='readonly',width=44);self.sortcol.pack(side='left',padx=5)
        ttk.Button(second,text='↑',command=lambda:self.sort(False)).pack(side='left');ttk.Button(second,text='↓',command=lambda:self.sort(True)).pack(side='left')
        ttk.Separator(second,orient='vertical').pack(side='left',fill='y',padx=9)
        ttk.Label(second,text='Столбец:').pack(side='left')
        ttk.Button(second,text='Авто',width=7,command=self.grid_auto_fit).pack(side='left',padx=2)
        self.count=ttk.Label(second,text='');self.count.pack(side='right')
        self.formula=tk.StringVar(value='Выберите ячейку — здесь будет полный текст. Ctrl+C / ⌘C копирует выделенный диапазон.')
        ttk.Entry(self,textvariable=self.formula,state='readonly').pack(fill='x',pady=(0,5))
        self.grid=ExcelGrid(self,on_select=self.formula.set,on_filter=self.open_filter);self.grid.pack(fill='both',expand=True)
    def grid_auto_fit(self):self.grid.auto_fit_toggle()
    def change_view(self,event=None):
        self.selected_product=None;self.column_filters={};self.dismiss_suggestions();self.grid.auto_restore.clear();self.grid.x=self.grid.y=0;self.refresh()
    def set_views(self,views):
        old=self.choice.get();self.views=views;self.column_filters={};self.selected_product=None;self.dismiss_suggestions();self.grid.auto_restore.clear();self.grid.x=self.grid.y=0;self.choice['values']=list(views)
        if old in views:self.choice.set(old)
        elif views:self.choice.current(0)
        else:self.choice.set('')
        self.refresh()
    def reset_search(self):
        self.query.set('');self.cancel_live_search();self.selected_product=None;self.column_filters={};self.dismiss_suggestions();self.refresh()
    def cancel_live_search(self):
        if self._pending_search is not None:
            self.after_cancel(self._pending_search);self._pending_search=None
    def on_query_change(self,*_):
        if self._choosing_suggestion:return
        self.selected_product=None
        if self._pending_search is not None:self.after_cancel(self._pending_search)
        self._pending_search=self.after(80,self.apply_live_search)
    def apply_live_search(self):
        self._pending_search=None
        self.refresh()
        if self.focus_get()==self.search_entry:self.show_suggestions()
    def show_suggestions(self):
        raw=self.views.get(self.choice.get(),pd.DataFrame())
        found=product_suggestions(raw,self.query.get())
        if self.suggest_popup is not None and found==self.suggestions:
            return
        self.dismiss_suggestions()
        self.suggestions=found
        if not self.suggestions:return
        self.suggest_popup=tk.Toplevel(self)
        self.suggest_popup.overrideredirect(True);self.suggest_popup.attributes('-topmost',True)
        box=tk.Listbox(self.suggest_popup,height=len(self.suggestions),width=70,activestyle='none',
                       bg='white',fg='#123149',selectbackground='#079bea',font=('Arial',10))
        box.pack(fill='both',expand=True)
        for title,_ in self.suggestions:box.insert('end',title)
        self.search_entry.update_idletasks()
        self.suggest_popup.geometry(f'+{self.search_entry.winfo_rootx()}+{self.search_entry.winfo_rooty()+self.search_entry.winfo_height()}')
        box.bind('<ButtonRelease-1>',lambda e:self.choose_suggestion(box))
    def dismiss_suggestions(self):
        if self.suggest_popup is not None:
            self.suggest_popup.destroy();self.suggest_popup=None
    def choose_suggestion(self,box):
        chosen=box.curselection()
        if chosen:
            self.cancel_live_search()
            title,self.selected_product=self.suggestions[chosen[0]]
            self._choosing_suggestion=True
            try:self.query.set(title)
            finally:self._choosing_suggestion=False
            self.dismiss_suggestions();self.refresh()
    def search_submit(self,event=None):
        self.cancel_live_search()
        if not self.query.get().strip():
            self.show_suggestions();return
        self.dismiss_suggestions();self.refresh()
    def refresh(self):
        df=self.views.get(self.choice.get(),pd.DataFrame()).copy();total_rows=len(df);q=self.query.get().strip()
        if self.selected_product is not None and not df.empty:df=df[exact_product(df,self.selected_product)]
        elif q and not df.empty:df=df[product_matches(df,q)]
        if self.hide.get() and not q and not df.empty and not self.choice.get().startswith('Исходные'):
            candidates=[c for c in df if '|' in str(c) and not any(s in str(c).lower() for s in ('артикул','sku','цена','дней','дн','срок','дата','строка','процент','доля','уровень','рейтинг','наименование','название','индекс','источник','статус','схема','штрих','категори'))]
            numeric=[c for c in candidates if pd.to_numeric(df[c].dropna(),errors='coerce').notna().all()]
            if numeric:
                active=df[numeric].apply(pd.to_numeric,errors='coerce').fillna(0).ne(0).any(axis=1)
                for c in df:
                    if 'Рекомендац' in str(c):active|=df[c].fillna('').astype(str).str.strip().ne('')
                df=df[active]
        df=present(df,self.show_calculations.get())
        df=apply_column_filters(df,self.column_filters)
        shown=visible(df,self.hide.get())
        product_fields=('Артикул','SKU','Название товара')
        has_product=any(df[k].fillna('').astype(str).str.strip().ne('').any() for k in product_fields if k in df)
        keep=[col for col in df if col in shown or self.column_filters.get(col) is not None
              or (has_product and col in product_fields)]
        self.filtered=df.loc[:,keep]
        self.sortcol['values']=list(self.filtered.columns)
        if len(self.filtered.columns):self.sortcol.current(0)
        else:self.sortcol.set('')
        self.grid.set_data(self.filtered);self.grid.set_filter_columns(self.column_filters)
        self.count.config(text=f'Строк: {len(self.filtered)} из {total_rows} · Столбцов: {len(self.filtered.columns)}')
    def open_filter(self,column,x,y):
        raw=self.views.get(self.choice.get(),pd.DataFrame())
        data=present(raw,self.show_calculations.get())
        if column not in data:return
        values=sorted(set(data[column].map(display)),key=lambda s:s.casefold())
        current=self.column_filters.get(column)
        checked=set(values) if current is None else set(current)
        win=tk.Toplevel(self);win.title(f'Фильтр: {column}');win.geometry(f'330x400+{x}+{y}');win.transient(self.winfo_toplevel())
        ttk.Label(win,text=column,padding=8).pack(fill='x')
        search=tk.StringVar();entry=ttk.Entry(win,textvariable=search);entry.pack(fill='x',padx=8,pady=4)
        body=ttk.Frame(win);body.pack(fill='both',expand=True,padx=8)
        listing=tk.Listbox(body,selectmode='browse',activestyle='none',bg='white',fg='#123149',font=('Arial',10))
        listing.pack(side='left',fill='both',expand=True)
        scroll=ttk.Scrollbar(body,command=listing.yview);scroll.pack(side='right',fill='y');listing.config(yscrollcommand=scroll.set)
        shown=[];notice=ttk.Label(win,text='');notice.pack(fill='x',padx=8)
        def render(*_):
            matching=[v for v in values if search.get().casefold() in v.casefold()]
            shown[:]=matching[:300];listing.delete(0,'end')
            for value in shown:listing.insert('end',('☑  ' if value in checked else '☐  ')+(value if value else '(пусто)'))
            notice.config(text=f'Показано {len(shown)} из {len(matching)} значений. Поиск уточняет список.')
        def toggle(event):
            idx=listing.nearest(event.y)
            if not 0<=idx<len(shown):return
            value=shown[idx]
            if value in checked:checked.remove(value)
            else:checked.add(value)
            listing.delete(idx);listing.insert(idx,('☑  ' if value in checked else '☐  ')+(value if value else '(пусто)'))
        listing.bind('<ButtonRelease-1>',toggle);search.trace_add('write',render);render()
        actions=ttk.Frame(win);actions.pack(fill='x',padx=8,pady=8)
        ttk.Button(actions,text='Все',command=lambda:(checked.update(values),render())).pack(side='left')
        ttk.Button(actions,text='Снять',command=lambda:(checked.clear(),render())).pack(side='left',padx=4)
        def apply():
            self.column_filters[column]=None if len(checked)==len(values) else checked
            win.destroy();self.refresh()
        ttk.Button(actions,text='Применить',command=apply).pack(side='right')
        entry.focus_set()
    def sort(self,descending):
        c=self.sortcol.get()
        if c not in self.filtered:return
        n=pd.to_numeric(self.filtered[c],errors='coerce')
        key=n if n.notna().sum()==self.filtered[c].notna().sum() else self.filtered[c].fillna('').astype(str).str.lower()
        order=key.sort_values(ascending=not descending,kind='stable',na_position='last').index
        self.filtered=self.filtered.loc[order].reset_index(drop=True);self.grid.set_data(self.filtered)
    def export(self):
        if self.filtered.empty:return
        path=filedialog.asksaveasfilename(initialfile='Глазок.xlsx',defaultextension='.xlsx',filetypes=[('Excel','*.xlsx'),('CSV','*.csv')])
        if not path:return
        try:
            if Path(path).suffix.lower()=='.csv':self.filtered.to_csv(path,index=False,sep=';',encoding='utf-8-sig')
            else:
                from openpyxl import Workbook
                from openpyxl.styles import Font,PatternFill,Border,Side,Alignment
                from openpyxl.utils import get_column_letter
                w=Workbook();ws=w.active;ws.title='Данные'
                ws.append([str(c) for c in self.filtered.columns])
                for row in self.filtered.itertuples(index=False,name=None):ws.append([None if pd.isna(v) else v for v in row])
                thin=Side(style='thin',color='D0D5DA')
                for row in ws:
                    for cell in row:
                        cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
                        if isinstance(cell.value,str):cell.data_type='s'
                for col_idx,cell in enumerate(ws[1],1):
                    calculated=str(self.filtered.columns[col_idx-1]).startswith('Расчёт | ')
                    cell.font=Font(bold=True,color='FFFFFF' if calculated else '123149')
                    cell.fill=PatternFill('solid',fgColor='7352A6' if calculated else 'DFF4FF')
                    cell.alignment=Alignment(wrap_text=True)
                    if calculated:
                        for row_idx in range(2,ws.max_row+1):
                            ws.cell(row_idx,col_idx).fill=PatternFill('solid',fgColor='F4EDFF')
                ws.row_dimensions[1].height=50;ws.freeze_panes='C2';ws.auto_filter.ref=ws.dimensions
                for i in range(1,len(self.filtered.columns)+1):ws.column_dimensions[get_column_letter(i)].width=28
                w.save(path)
            messagebox.showinfo('Сохранено','Сохранены все строки текущего представления с учётом поиска и скрытия нулей.')
        except Exception as e:messagebox.showerror('Экспорт',str(e))


class Glasok:
    def __init__(self,root):
        self.root=root;root.title('Глазок');root.geometry('1400x820');root.minsize(1050,600)
        self.paths=[];self.busy=False;self.messages=queue.Queue();self.panels={};self.manual_result={};self.api_views={'Ozon':{},'WB':{}};self.last_warnings=[]
        self._update_poll_started=False;self._offered_update=None
        set_app_icon(root)
        style=ttk.Style()
        if 'clam' in style.theme_names():style.theme_use('clam')
        root.configure(background='#dff4ff')
        style.configure('.',font=('Segoe UI',10),background='#dff4ff',foreground='#123149')
        style.configure('TFrame',background='#dff4ff')
        style.configure('TLabelframe',background='#dff4ff',bordercolor='#43b9f5',relief='solid')
        style.configure('TLabelframe.Label',background='#dff4ff',foreground='#0077b6',font=('Segoe UI',10,'bold'))
        style.configure('TLabel',background='#dff4ff',foreground='#123149')
        style.configure('TButton',background='#079bea',foreground='white',padding=(10,6),borderwidth=1,focuscolor='#005f99')
        style.map('TButton',background=[('active','#28b5ff'),('pressed','#007fc4'),('disabled','#9bd7f7')],foreground=[('disabled','#eefaff')])
        style.configure('TCheckbutton',background='#dff4ff',foreground='#123149')
        style.map('TCheckbutton',background=[('active','#dff4ff')],indicatorcolor=[('selected','#079bea'),('!selected','white')])
        style.configure('TNotebook',background='#dff4ff',borderwidth=0)
        style.configure('TNotebook.Tab',background='#a9e1ff',foreground='#075c8d',padding=(22,9),font=('Segoe UI',10,'bold'))
        style.map('TNotebook.Tab',background=[('selected','#079bea'),('active','#63c7ff')],foreground=[('selected','white'),('active','#073b5a')])
        style.configure('TEntry',fieldbackground='white',foreground='#123149',bordercolor='#43b9f5',lightcolor='#43b9f5',darkcolor='#43b9f5')
        style.configure('TCombobox',fieldbackground='white',background='white',foreground='#123149',arrowcolor='#007fc4',bordercolor='#43b9f5')
        style.map('TCombobox',fieldbackground=[('readonly','white')],selectbackground=[('readonly','white')],selectforeground=[('readonly','#123149')])
        style.configure('TScrollbar',background='#63c7ff',troughcolor='#d9f2ff',arrowcolor='#006da8',bordercolor='#d9f2ff')
        style.configure('TSeparator',background='#63c7ff')
        top=ttk.Frame(root,padding=10);top.pack(fill='x')
        self.addbutton=ttk.Button(top,text='Добавить файлы Ozon / WB',command=self.add);self.addbutton.pack(side='left')
        self.folderbutton=ttk.Button(top,text='Добавить папку',command=self.folder);self.folderbutton.pack(side='left',padx=6)
        self.clearbutton=ttk.Button(top,text='Очистить всё',command=self.clear);self.clearbutton.pack(side='left')
        ttk.Button(top,text='ⓘ Справка',command=self.help).pack(side='right')
        ttk.Button(top,text='Подключение к серверу',command=self.api_settings).pack(side='right',padx=7)
        self.server_state=ttk.Label(top,text='Сервер: проверка…');self.server_state.pack(side='right',padx=12)
        ttk.Label(root,text='Маркетплейс определяется по содержимому. Фильтры ▼ находятся в каждом столбце; ширина и высота меняются мышью за границы заголовков.',padding=(10,0,10,5),foreground='#006da8').pack(anchor='w')
        self.book=ttk.Notebook(root);self.book.pack(fill='both',expand=True,padx=10,pady=5)
        for name in ('Ozon','WB'):
            panel=ReportPanel(self.book,name,self.fetch_api)
            self.panels[name]=panel;self.book.add(panel,text=name)
        self.status=ttk.Label(root,text='Загрузите файлы или получите выбранные отчёты через API.',padding=10);self.status.pack(fill='x')
        threading.Thread(target=self.check_server,daemon=True).start()
    def check_server(self):
        ready=ensure_server()
        if ready:
            try:
                configured=api_request('/settings',timeout=3)
                markets=[name for name,key in (('Ozon','ozon'),('WB','wb')) if configured.get(key)]
                label='Сервер: доступен · ключи '+(', '.join(markets) if markets else 'не настроены')
                color='#18794e' if markets else '#a15c00'
            except Exception:
                label='Сервер: доступен · проверьте код подключения';color='#b42318'
        else:
            label='Сервер: не запущен';color='#b42318'
        self.root.after(0,lambda:self.server_state.config(text=label,foreground=color))
        self.root.after(10000,lambda:threading.Thread(target=self.check_server,daemon=True).start())
        if not self._update_poll_started:
            self._update_poll_started=True;self.root.after(15000,self.schedule_update_check)
    def add(self):
        if self.busy:return
        selected=filedialog.askopenfilenames(title='Все отчёты — Ozon и Wildberries вместе',filetypes=[('Таблицы и архивы','*.xlsx *.XLSX *.xlsm *.XLSM *.xls *.XLS *.csv *.CSV *.tsv *.TSV *.zip *.ZIP'),('Все файлы','*')])
        if selected:self.load(selected)
    def folder(self):
        if self.busy:return
        path=filedialog.askdirectory(title='Папка с отчётами (включая вложенные папки)')
        if path:
            files=[str(p) for p in Path(path).rglob('*') if p.is_file() and p.suffix.lower() in SUPPORTED and not p.name.startswith(('~$','.'))]
            if files:self.load(files)
            else:messagebox.showinfo('Папка','Поддерживаемых файлов не найдено.')
    def clear(self):
        if self.busy:return
        self.paths=[];self.manual_result={};self.api_views={'Ozon':{},'WB':{}};self.last_warnings=[]
        for p in self.panels.values():p.set_views({})
        self.remove_unknown_panel()
        self.status.config(text='Данные очищены')
    def remove_unknown_panel(self):
        panel=self.panels.pop('Не распознано',None)
        if panel is not None:self.book.forget(panel);panel.destroy()
    def load(self,paths):
        candidate=list(dict.fromkeys(self.paths+list(paths)));self.busy=True
        for button in (self.addbutton,self.folderbutton,self.clearbutton):button.config(state='disabled')
        def work():
            try:self.messages.put(('done',analyze_all(candidate,lambda t:self.messages.put(('progress',t)))))
            except Exception as e:self.messages.put(('error',str(e)))
        threading.Thread(target=work,daemon=True).start()
        def poll():
            finished=False
            while not self.messages.empty():
                kind,value=self.messages.get()
                if kind=='progress':self.status.config(text=value)
                else:
                    finished=True;self.busy=False
                    for button in (self.addbutton,self.folderbutton,self.clearbutton):button.config(state='normal')
                    if kind=='error':messagebox.showerror('Загрузка',value);self.status.config(text='Ошибка загрузки — предыдущие данные сохранены');break
                    result,warnings=value;self.paths=candidate;self.manual_result=result;self.last_warnings=warnings
                    for name in ('Ozon','WB'):
                        self.panels[name].set_views(self.merged_views(name))
                    unknown=result.get('Не распознано',{})
                    if unknown:
                        if 'Не распознано' not in self.panels:
                            panel=ReportPanel(self.book);self.panels['Не распознано']=panel
                            self.book.add(panel,text='Не распознано')
                        self.panels['Не распознано'].set_views(unknown)
                    else:self.remove_unknown_panel()
                    self.status.config(text=f'Выбрано файлов: {len(candidate)}. Замечаний: {len(warnings)}.'+(' Откройте ⓘ для подробностей.' if warnings else ''))
                    if warnings:
                        lines='\n'.join(str(item) for item in warnings[:12])
                        if len(warnings)>12:lines+=f'\n…и ещё {len(warnings)-12}. Полный список — в ⓘ Справка.'
                        messagebox.showwarning('Замечания при загрузке',lines)
                    if not result.get('Ozon') and result.get('WB'):self.book.select(self.panels['WB'])
            if not finished:self.root.after(100,poll)
        self.root.after(100,poll)

    def merged_views(self,market):
        manual=self.manual_result.get(market,{}) if self.manual_result else {};api=self.api_views.get(market,{})
        out={**manual,**api}
        frames=[]
        if not manual.get('Общая сводка товаров',pd.DataFrame()).empty:
            frames.append(manual['Общая сводка товаров'].copy())
        api_frame=api.get('API — Объединённая сводка',pd.DataFrame())
        if isinstance(api_frame,pd.DataFrame) and not api_frame.empty:
            frames.append(api_frame.copy())
        if frames:out={'Все источники вместе':pd.concat(frames,ignore_index=True,sort=False),**out}
        if market=='Ozon' and 'Продажи и текущие остатки' in out:
            out={'Продажи и текущие остатки':out['Продажи и текущие остатки'],**out}
        return out

    def fetch_api(self,market,reports,start,end):
        if self.busy:return
        if not reports:return messagebox.showinfo('API','Отметьте хотя бы один отчёт.')
        try:
            date.fromisoformat(start);date.fromisoformat(end)
            if start>end:raise ValueError
        except ValueError:return messagebox.showerror('Период','Введите даты в формате ГГГГ-ММ-ДД. Начало должно быть не позже конца.')
        self.busy=True;self.status.config(text=f'{market}: сервер получает выбранные отчёты…')
        for panel in self.panels.values():
            if hasattr(panel,'api_button'):panel.api_button.config(state='disabled')
        def work():
            try:
                if not ensure_server():raise RuntimeError('Не удалось запустить Глазок Сервер')
                value=api_request('/fetch',{'marketplace':market,'reports':reports,'date_from':start,'date_to':end})
                self.messages.put(('api_done',(market,value)))
            except Exception as exc:self.messages.put(('api_error',str(exc)))
        threading.Thread(target=work,daemon=True).start();self.root.after(100,self.poll_api)

    def poll_api(self):
        finished=False
        while not self.messages.empty():
            kind,value=self.messages.get()
            if kind not in ('api_done','api_error'):
                self.messages.put((kind,value));break
            finished=True;self.busy=False
            for panel in self.panels.values():
                if hasattr(panel,'api_button'):panel.api_button.config(state='normal')
            if kind=='api_error':messagebox.showerror('API',value);self.status.config(text='Ошибка API — прежние данные сохранены');continue
            market,payload=value
            converted={name:pd.DataFrame(rows) for name,rows in payload.get('views',{}).items()}
            self.api_views[market].update(converted);self.panels[market].set_views(self.merged_views(market));self.book.select(self.panels[market])
            errors=payload.get('errors') or []
            self.status.config(text=f'{market}: получено представлений {len(converted)}. Ошибок отдельных отчётов: {len(errors)}.')
            if errors:
                detail=('Проблема с сертификатом HTTPS затронула '+str(len(errors))+' отчётов.\n\n'+errors[0].split(': ',1)[-1]
                        if all('Не удалось проверить сертификат HTTPS' in error for error in errors)
                        else '\n\n'.join(errors))
                messagebox.showwarning('Часть API-отчётов не получена',detail)
        if not finished:self.root.after(100,self.poll_api)

    def api_settings(self):
        win=tk.Toplevel(self.root);win.title('Подключение к серверу');win.resizable(False,False);win.transient(self.root);win.grab_set()
        box=ttk.Frame(win,padding=14);box.pack(fill='both',expand=True)
        current,token=connection();address=tk.StringVar(value=current);secret=tk.StringVar(value=token if current!=BASE else '')
        ttk.Label(box,text='Адрес сервера').grid(row=0,column=0,sticky='w',pady=5)
        ttk.Entry(box,textvariable=address,width=52).grid(row=0,column=1,padx=12,pady=5)
        ttk.Label(box,text='Код подключения').grid(row=1,column=0,sticky='w',pady=5)
        ttk.Entry(box,textvariable=secret,width=52,show='•').grid(row=1,column=1,padx=12,pady=5)
        ttk.Label(box,text='Адрес и код показаны в окне «Глазок Сервер».\nAPI-ключи Ozon и WB вводятся только там.',foreground='#555').grid(row=2,column=0,columnspan=2,sticky='w',pady=(10,8))
        def save():
            try:
                from urllib.request import Request,urlopen
                target=address.get().strip().rstrip('/')
                with urlopen(Request(target+'/catalog',headers={'Authorization':'Bearer '+(secret.get().strip() or token)}),timeout=4):
                    pass
                save_connection(address.get(),secret.get() or token)
                win.destroy();self.check_server()
            except Exception as exc:messagebox.showerror('Подключение',str(exc),parent=win)
        ttk.Button(box,text='Сохранить и проверить',command=save).grid(row=3,column=1,sticky='e')
    def schedule_update_check(self):
        if sys.platform!='win32' or not getattr(sys,'frozen',False):return
        def work():
            try:
                item=api_request('/updates/latest',timeout=5).get('update')
                if item:
                    staged=download_update(item)
                    if staged and self._offered_update!=item['sha256']:
                        self._offered_update=item['sha256']
                        self.root.after(0,lambda:self.offer_update(staged))
            except Exception as exc:
                detail=str(exc)
                self.root.after(0,lambda:self.status.config(text=f'Проверка обновлений: {detail}'))
            finally:self.root.after(60000,self.schedule_update_check)
        threading.Thread(target=work,daemon=True).start()
    def offer_update(self,staged):
        if messagebox.askyesno('Обновление Глазок','Новая версия загружена. Перезапустить приложение сейчас?'):
            try:launch_update(staged)
            except Exception as exc:return messagebox.showerror('Обновление',str(exc))
            self.root.destroy()
    def help(self):
        win=tk.Toplevel(self.root);win.title('Глазок — информация');win.geometry('840x680');win.minsize(600,420)
        body=ttk.Frame(win,padding=12);body.pack(fill='both',expand=True)
        text=tk.Text(body,wrap='word',bg='#ffffff',fg='#123149',font=('Arial',11),padx=16,pady=12)
        scroll=ttk.Scrollbar(body,orient='vertical',command=text.yview)
        text.configure(yscrollcommand=scroll.set);text.pack(side='left',fill='both',expand=True)
        scroll.pack(side='right',fill='y')
        content=HELP_TEXT
        if self.last_warnings:
            content+='\n\nЗАМЕЧАНИЯ ПО ПОСЛЕДНЕЙ ЗАГРУЗКЕ\n'+'\n'.join('• '+str(w) for w in self.last_warnings)
        text.insert('1.0',content);text.configure(state='disabled')


def main():
    if sys.platform=='win32':
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('Aristocrat.Glazok.5')
    root=tk.Tk();Glasok(root);root.mainloop()

if __name__=='__main__':main()
