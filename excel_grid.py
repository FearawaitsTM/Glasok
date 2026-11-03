"""Виртуализированная таблица Tk: рисуются только видимые ячейки."""
import bisect
import csv
import io
import math
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk
import pandas as pd


def display(v):
    if v is None:return ''
    try:
        if pd.isna(v):return ''
    except (TypeError,ValueError):pass
    if isinstance(v,float):return f'{v:,.4f}'.rstrip('0').rstrip('.').replace(',',' ')
    return str(v)


def letter(n):
    out='';n+=1
    while n:n,r=divmod(n-1,26);out=chr(65+r)+out
    return out


class ExcelGrid(ttk.Frame):
    ROW=26;HEADER=66;INDEX=60
    LETTER_HEIGHT=22
    def __init__(self,parent,on_select=lambda text:None,on_filter=lambda col,x,y:None):
        super().__init__(parent);self.on_select=on_select;self.on_filter=on_filter;self.df=pd.DataFrame();self.values=[];self.widths=[];self.saved_widths={};self.auto_restore={};self.edges=[0];self.row_heights=[];self.row_edges=[0];self.header_height=self.HEADER;self.index_width=self.INDEX;self.x=0;self.y=0;self.anchor=None;self.end=None;self.resize_kind=None;self.active_filters=set()
        self.cell_font=tkfont.Font(family='Arial',size=10);self.header_font=tkfont.Font(family='Arial',size=9);self.letter_font=tkfont.Font(family='Arial',size=9,weight='bold')
        self.canvas=tk.Canvas(self,bg='#f7fcff',highlightthickness=1,highlightbackground='#63c7ff',takefocus=True)
        self.vertical=ttk.Scrollbar(self,orient='vertical',command=self.yscroll);self.horizontal=ttk.Scrollbar(self,orient='horizontal',command=self.xscroll)
        self.canvas.grid(row=0,column=0,sticky='nsew');self.vertical.grid(row=0,column=1,sticky='ns');self.horizontal.grid(row=1,column=0,sticky='ew');self.rowconfigure(0,weight=1);self.columnconfigure(0,weight=1)
        self.canvas.bind('<Configure>',lambda e:self.draw());self.canvas.bind('<Button-1>',self.click);self.canvas.bind('<B1-Motion>',self.drag);self.canvas.bind('<ButtonRelease-1>',lambda e:setattr(self,'resize_kind',None))
        self.canvas.bind('<Motion>',self.hover);self.canvas.bind('<Leave>',lambda e:self.canvas.config(cursor=''))
        self.canvas.bind('<Double-Button-1>',self.double_click)
        self.canvas.bind('<MouseWheel>',self.wheel);self.canvas.bind('<Shift-MouseWheel>',self.wheel)
        self.canvas.bind('<Button-4>',lambda e:self.yscroll('scroll',-3,'units'));self.canvas.bind('<Button-5>',lambda e:self.yscroll('scroll',3,'units'))
        for event in ('<Control-c>','<Command-c>'):self.canvas.bind(event,self.copy)
        for event in ('<Control-a>','<Command-a>'):self.canvas.bind(event,self.select_all)
        for key in ('Up','Down','Left','Right','Home','End','Next','Prior'):self.canvas.bind('<'+key+'>',self.key)
        self.menu=tk.Menu(self,tearoff=False);self.menu.add_command(label='Копировать ячейки',command=self.copy);self.menu.add_command(label='Выделить всё',command=self.select_all)
        self.canvas.bind('<Button-3>',lambda e:self.menu.tk_popup(e.x_root,e.y_root))
    def set_data(self,df):
        self.df=df.reset_index(drop=True);self.values=self.df.values.tolist()
        self.widths=[self.saved_widths.get(str(c),280 if any(t in str(c) for t in ('Название','Наименование','Источник','Причина','Сообщение')) else 205) for c in df.columns]
        if len(self.row_heights)!=len(self.values):self.row_heights=[self.ROW]*len(self.values)
        self.anchor=self.end=None;self.layout();self.draw()
    def set_filter_columns(self,filters):
        self.active_filters={c for c,v in filters.items() if v is not None};self.draw()
    def layout(self):
        self.edges=[0]
        for w in self.widths:self.edges.append(self.edges[-1]+w)
        self.row_edges=[0]
        for h in self.row_heights:self.row_edges.append(self.row_edges[-1]+h)
    def limits(self):return max(1,self.canvas.winfo_width()-self.index_width),max(1,self.canvas.winfo_height()-self.header_height)
    def bounds(self):
        if self.anchor is None or self.end is None:return None
        return min(self.anchor[0],self.end[0]),max(self.anchor[0],self.end[0]),min(self.anchor[1],self.end[1]),max(self.anchor[1],self.end[1])
    def draw(self):
        c=self.canvas;c.delete('all');vw,vh=self.limits();totalw=self.edges[-1];totalh=self.row_edges[-1]
        self.x=max(0,min(self.x,max(0,totalw-vw)));self.y=max(0,min(self.y,max(0,totalh-vh)))
        self.horizontal.set(self.x/max(1,totalw),min(1,(self.x+vw)/max(1,totalw)));self.vertical.set(self.y/max(1,totalh),min(1,(self.y+vh)/max(1,totalh)))
        startcol=max(0,bisect.bisect_right(self.edges,self.x)-1);endcol=min(len(self.widths),bisect.bisect_left(self.edges,self.x+vw)+1)
        first=max(0,bisect.bisect_right(self.row_edges,self.y)-1);last=min(len(self.values),bisect.bisect_right(self.row_edges,self.y+vh)+1);selection=self.bounds()
        for r in range(first,last):
            y=self.header_height+self.row_edges[r]-self.y;h=self.row_heights[r]
            for col in range(startcol,endcol):
                x=self.index_width+self.edges[col]-self.x;w=self.widths[col]
                calculated=str(self.df.columns[col]).startswith('Расчёт | ')
                selected=selection and selection[0]<=r<=selection[1] and selection[2]<=col<=selection[3]
                background=('#ddd0f7' if selected else ('#f4edff' if r%2 else '#faf6ff')) if calculated else ('#cceeff' if selected else ('#f7fcff' if r%2 else 'white'))
                c.create_rectangle(x,y,x+w,y+h,fill=background,outline='#b6def5',width=1)
                text=display(self.values[r][col]).replace('\n',' ↵ ')
                maxchars=max(4,int((w-10)/7.4));text=text if len(text)<=maxchars else text[:maxchars-1]+'…'
                c.create_text(x+5,y+h/2,anchor='w',text=text,font=self.cell_font,fill='#123149')
        # Неподвижные заголовки и номера строк перекрывают прокрученное содержимое.
        c.create_rectangle(0,0,vw+self.index_width,self.header_height,fill='#079bea',outline='')
        for col in range(startcol,endcol):
            x=self.index_width+self.edges[col]-self.x;w=self.widths[col]
            calculated=str(self.df.columns[col]).startswith('Расчёт | ')
            c.create_rectangle(x,0,x+w,self.header_height,fill='#7352a6' if calculated else '#079bea',outline='#caecff',width=1)
            title=str(self.df.columns[col]);maxchars=max(4,int((w-32)/7))
            words=title.split();lines=[];line=''
            for word in words:
                if len(line)+len(word)>maxchars and line:lines.append(line);line=word
                else:line+=(' ' if line else '')+word
            if line:lines.append(line)
            shown='\n'.join(lines[:3]);shown+=('…' if len(lines)>3 else '')
            c.create_text(x+w/2,10,text=letter(col),font=self.letter_font,fill='#eaf8ff')
            c.create_text(x+5,self.LETTER_HEIGHT+2,anchor='nw',text=shown,font=self.header_font,fill='white')
            c.create_rectangle(x+w-24,self.LETTER_HEIGHT+2,x+w-4,self.header_height-5,
                               fill='#0076ba' if title in self.active_filters else '#36b5f5',outline='#d7f1ff')
            c.create_text(x+w-14,(self.LETTER_HEIGHT+self.header_height)/2,text='▼',font=('Arial',8),fill='white')
            # Яркая полоса показывает, за что тянуть для изменения ширины.
            c.create_line(x+w-1,0,x+w-1,self.header_height,fill='#ffffff',width=2)
            c.create_line(x+w-4,4,x+w-4,self.LETTER_HEIGHT-4,fill='#bce9ff',width=1)
        c.create_rectangle(0,self.header_height,self.index_width,vh+self.header_height,fill='#d9f2ff',outline='')
        for r in range(first,last):
            y=self.header_height+self.row_edges[r]-self.y;h=self.row_heights[r]
            c.create_rectangle(0,y,self.index_width,y+h,fill='#d9f2ff',outline='#9bd7f7');c.create_text(self.index_width-8,y+h/2,anchor='e',text=str(r+1),font=self.cell_font,fill='#075c8d')
        c.create_rectangle(0,0,self.index_width,self.header_height,fill='#007fc4',outline='#caecff')
        c.create_text(self.index_width/2,self.header_height/2,text='☰',font=('Arial',16,'bold'),fill='white')
        if not self.values:c.create_text(self.index_width+20,self.header_height+28,anchor='nw',text='Нет строк для отображения. Измените фильтр ▼ или нажмите «Сбросить».',fill='#075c8d')
    def coords(self,event):
        if self.df.empty:return None
        row=bisect.bisect_right(self.row_edges,event.y-self.header_height+self.y)-1;col=bisect.bisect_right(self.edges,event.x-self.index_width+self.x)-1
        return max(0,min(row,len(self.values)-1)),max(0,min(col,len(self.widths)-1))
    def click(self,event):
        self.canvas.focus_set()
        if not self.widths:return
        if event.y>=self.LETTER_HEIGHT and event.y<self.header_height-6 and event.x>=self.index_width:
            col=max(0,min(bisect.bisect_right(self.edges,event.x-self.index_width+self.x)-1,len(self.widths)-1))
            if self.index_width+self.edges[col]-self.x<=event.x and event.x>=self.index_width+self.edges[col+1]-self.x-26:
                self.on_filter(str(self.df.columns[col]),event.x_root,event.y_root);return
        if abs(event.x-self.index_width)<6:
            self.resize_kind='index';self.resize_start=(event.x,self.index_width);return
        if event.y<self.LETTER_HEIGHT:
            col=self.boundary_at(event.x)
            if col is not None:self.resize_kind='column';self.resize_target=col;self.resize_start=(event.x,self.widths[col]);return
        if abs(event.y-self.header_height)<6 and event.x>=self.index_width:
            self.resize_kind='header';self.resize_start=(event.y,self.header_height);return
        if not self.values:return
        if event.x<self.index_width and event.y>=self.header_height:
            row=self.row_boundary_at(event.y)
            if row is not None:self.resize_kind='row';self.resize_target=row;self.resize_start=(event.y,self.row_heights[row]);return
        if event.y<self.header_height:
            _,col=self.coords(event)
            self.anchor=(0,col);self.end=(len(self.values)-1,col)
        elif event.x<self.index_width:
            row,_=self.coords(event);self.anchor=(row,0);self.end=(row,len(self.widths)-1)
        else:
            pos=self.coords(event)
            if not event.state&1 or self.anchor is None:self.anchor=pos
            self.end=pos
        self.describe();self.draw()
    def drag(self,event):
        if self.resize_kind=='column':
            self.set_width(self.resize_target,self.resize_start[1]+event.x-self.resize_start[0]);return
        if self.resize_kind=='row':
            row=self.resize_target;self.row_heights[row]=max(20,min(180,self.resize_start[1]+event.y-self.resize_start[0]));self.layout();self.draw();return
        if self.resize_kind=='header':
            self.header_height=max(48,min(200,self.resize_start[1]+event.y-self.resize_start[0]));self.draw();return
        if self.resize_kind=='index':
            self.index_width=max(42,min(180,self.resize_start[1]+event.x-self.resize_start[0]));self.draw();return
        if self.anchor is not None:self.end=self.coords(event);self.describe();self.draw()
    def boundary_at(self,x):
        for col,edge in enumerate(self.edges[1:]):
            if abs(x-(edge+self.index_width-self.x))<7:return col
        return None
    def row_boundary_at(self,y):
        edge=y-self.header_height+self.y
        row=bisect.bisect_left(self.row_edges,edge)
        for candidate in (row-1,row):
            if 1<=candidate<len(self.row_edges) and abs(edge-self.row_edges[candidate])<6:return candidate-1
        return None
    def hover(self,event):
        horizontal=abs(event.x-self.index_width)<6 or (event.y<self.LETTER_HEIGHT and self.boundary_at(event.x) is not None)
        vertical=(abs(event.y-self.header_height)<6 and event.x>=self.index_width) or (event.x<self.index_width and event.y>=self.header_height and self.row_boundary_at(event.y) is not None)
        self.canvas.config(cursor='sb_h_double_arrow' if horizontal else 'sb_v_double_arrow' if vertical else '')
    def double_click(self,event):
        if event.y<self.LETTER_HEIGHT:
            col=self.boundary_at(event.x)
            if col is not None:self.auto_fit_toggle(col);return 'break'
    def selected_col(self):
        if self.end is not None:return self.end[1]
        if self.anchor is not None:return self.anchor[1]
        return 0 if self.widths else None
    def set_width(self,col,width,keep_auto=False):
        if col is None or not 0<=col<len(self.widths):return
        if not keep_auto:self.auto_restore.pop(str(self.df.columns[col]),None)
        self.widths[col]=max(65,min(1100,int(width)));self.saved_widths[str(self.df.columns[col])]=self.widths[col];self.layout();self.draw()
    def auto_fit_toggle(self,col=None):
        col=self.selected_col() if col is None else col
        if col is None:return
        name=str(self.df.columns[col])
        if name in self.auto_restore:
            previous,auto_width=self.auto_restore.pop(name)
            if self.widths[col]==auto_width:
                self.set_width(col,previous,keep_auto=True)
                return
        previous=self.widths[col]
        title=max((self.header_font.measure(part) for part in str(self.df.columns[col]).split()),default=0)+22
        sample=(display(row[col]).replace('\n',' ') for row in self.values[:1000])
        content=max((self.cell_font.measure(value) for value in sample),default=0)+18
        auto_width=max(80,min(850,max(title,content)))
        self.set_width(col,auto_width,keep_auto=True)
        self.auto_restore[name]=(previous,self.widths[col])
    def describe(self):
        if self.end is None:return
        r,c=self.end;bounds=self.bounds();count=(bounds[1]-bounds[0]+1)*(bounds[3]-bounds[2]+1)
        self.on_select(f'{letter(c)}{r+1}  |  {self.df.columns[c]}  |  {display(self.values[r][c])}  ·  Выбрано ячеек: {count}')
    def copy(self,event=None):
        b=self.bounds()
        if not b:return 'break'
        buf=io.StringIO();writer=csv.writer(buf,delimiter='\t',lineterminator='\n')
        for row in self.values[b[0]:b[1]+1]:writer.writerow([display(v) for v in row[b[2]:b[3]+1]])
        self.clipboard_clear();self.clipboard_append(buf.getvalue());return 'break'
    def select_all(self,event=None):
        if not self.df.empty:self.anchor=(0,0);self.end=(len(self.values)-1,len(self.widths)-1);self.describe();self.draw()
        return 'break'
    def yscroll(self,*args):
        total=self.row_edges[-1];_,vh=self.limits()
        if args[0]=='moveto':self.y=float(args[1])*total
        else:self.y+=int(args[1])*(vh if args[2]=='pages' else self.ROW)
        self.draw()
    def xscroll(self,*args):
        vw,_=self.limits()
        if args[0]=='moveto':self.x=float(args[1])*self.edges[-1]
        else:self.x+=int(args[1])*(vw if args[2]=='pages' else 40)
        self.draw()
    def wheel(self,event):
        delta=event.delta;step=-int(delta/120) if abs(delta)>=120 else (-1 if delta>0 else 1)
        if event.state&1:self.xscroll('scroll',step*2,'units')
        else:self.yscroll('scroll',step*3,'units')
        return 'break'
    def key(self,event):
        if self.df.empty:return 'break'
        r,c=self.end or (0,0);page=max(1,int(self.limits()[1]/self.ROW))
        dr,dc={'Up':(-1,0),'Down':(1,0),'Left':(0,-1),'Right':(0,1),'Home':(0,-c),'End':(0,len(self.widths)-1-c),'Next':(page,0),'Prior':(-page,0)}[event.keysym]
        r=max(0,min(len(self.values)-1,r+dr));c=max(0,min(len(self.widths)-1,c+dc));self.end=(r,c)
        if not event.state&1 or self.anchor is None:self.anchor=self.end
        vw,vh=self.limits()
        if self.row_edges[r]<self.y:self.y=self.row_edges[r]
        if self.row_edges[r+1]>self.y+vh:self.y=self.row_edges[r+1]-vh
        if self.edges[c]<self.x:self.x=self.edges[c]
        if self.edges[c+1]>self.x+vw:self.x=self.edges[c+1]-vw
        self.describe();self.draw();return 'break'
