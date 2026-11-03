"""Общая загрузка: содержимое определяет маркетплейс, расширение — способ чтения."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile,is_zipfile
from collections import defaultdict
import hashlib
import math
import pandas as pd
from marketplace_wb import load_workbook_safely,close_workbook,analyze_wb,wb_period,number_value
from ozon_engine import analyze_ozon,clean,num,ident

SUPPORTED={'.xlsx','.xlsm','.xls','.csv','.tsv','.zip'}


def classify(path):
    w=load_workbook_safely(path)
    try:
        kinds=set();preview={}
        for ws in w:
            rows=list(ws.values)
            for row in rows[:35]:
                h={clean(v).lower() for v in row if v is not None}
                if 'артикул wb' in h or '№ отчета' in h and 'к перечислению за товар' in h:kinds.add('WB')
                if ('sku' in h and ('артикул' in h or 'статус проверки' in h or 'номер поставки' in h or 'дата возврата' in h)) or ('уровень выкупа товара (за последние 50 заказов)' in h and 'артикул' in h):kinds.add('Ozon')
            if not kinds and rows:preview[ws.title]=rows
        if len(kinds)==1:return next(iter(kinds)),{}
        if len(kinds)>1:return 'Не определён: смешанные листы',preview
        return 'Не определён',preview
    finally:close_workbook(w)


def safe_sum(s):
    n=pd.to_numeric(s,errors='coerce');return n.sum(min_count=1)


def enrich_wb(df):
    if df.empty:return df
    records=df.to_dict('records')
    for r in records:
        p=wb_period(r.get('Период'));days=(p[1]-p[0]).days+1 if p else None
        def get(c):
            v=num(r.get(c));return v if v is not None and math.isfinite(v) else None
        def ratio(label,a,b,factor=1):
            if a is not None and b is not None and b>0:r['Расчёт | '+label]=a/b*factor
        ordered=get('Поставщик | шт.');redeemed=get('Поставщик | Выкупили, шт.');stock=get('Поставщик | Текущий остаток, шт.')
        ratio('Заказов в день — поставщик',ordered,days)
        ratio('Выкупов в день — поставщик',redeemed,days)
        ratio('Выкуплено / заказано — поставщик, %',redeemed,ordered,100)
        ratio('Среднее перечисление за выкуп, ₽',get('Поставщик | К перечислению за товар, руб.'),redeemed)
        # Текущий остаток старого отчёта не является остатком на конец прошлого месяца.
        if stock is not None:r['Поставщик | Примечание к остатку']='Текущий остаток из отчёта; не остаток на конец выбранного периода'
        sale=get('Доходы и расходы | Продажи, ₽');units=get('Доходы и расходы | Продажи, шт');result=get('Доходы и расходы | Итог, ₽')
        ratio('Средняя сумма продажи — Д и Р, ₽',sale,units)
        ratio('Итог / продажи — Д и Р, %',result,sale,100)
        ratio('Продано в день — Д и Р',units,days)
        costs=[get('Доходы и расходы | '+c) for c in ('Логистика, ₽','Штрафы, ₽','Комиссия WB, ₽','Эквайринг, ₽','Потери, подмены и товары с дефектами, ₽','Программа лояльности, ₽')]
        present=[x for x in costs if x is not None]
        if present:
            r['Расчёт | Удержания — Д и Р, ₽']=-sum(min(x,0) for x in present)
            ratio('Удержания / продажи — Д и Р, %',r['Расчёт | Удержания — Д и Р, ₽'],sale,100)
        if units is not None and redeemed is not None:r['Расчёт | Продажи Д и Р − выкупы поставщика, шт']=units-redeemed
        for prefix in ('Остатки: детали','Воронка'):
            q=get(prefix+' | Заказали, шт') if prefix=='Остатки: детали' else get(prefix+' | Заказали товаров, шт')
            b=get(prefix+' | Выкупили, шт')
            ratio('Заказов в день — '+prefix,q,days)
            ratio('Выкупов в день — '+prefix,b,days)
            ratio('Выкуплено / заказано — '+prefix+', %',b,q,100)
    return pd.DataFrame(records)


def join_period_views(full,daily,market):
    """Join только при совпадении полного интервала, не дробим агрегаты."""
    if full.empty:
        out=daily.copy()
        if not out.empty:out.insert(1,'Тип блока','Дневная детализация')
        return out
    if daily.empty:
        out=full.copy();out.insert(1,'Тип блока','Полный период');return out
    keys=['Период','Артикул WB'] if market=='WB' else ['Период','Артикул']
    # Ozon CSV не содержит схему. Присоединяем лишь когда схема у артикула
    # однозначна, иначе оставляем отдельный блок, не размножая значения CSV.
    f=full.copy();d=daily.copy()
    d=d.rename(columns={c:c+' — дневные данные' for c in d if c in f and '|' in c})
    for frame in (f,d):
        for k in keys:frame[k]=frame[k].fillna('').astype(str)
    invalid=d[keys[-1]].eq('')
    counts=f.groupby(keys,dropna=False).size()
    eligible=[not invalid.iloc[i] and counts.get(tuple(r),0)==1 for i,r in enumerate(d[keys].itertuples(index=False,name=None))]
    attach=d[eligible];remaining=d[[not v for v in eligible]]
    merged=f.merge(attach,on=keys,how='left',suffixes=('','__day'),validate='one_to_one' if not f.duplicated(keys).any() else 'many_to_one',indicator=True)
    merged.insert(1,'Тип блока',merged['_merge'].map({'both':'Полный период + дневные данные','left_only':'Полный период','right_only':'Дневная детализация'}).astype(str));merged=merged.drop(columns='_merge')
    for c in list(merged):
        if c.endswith('__day'):
            target=c[:-5]
            if target in merged:merged[target]=merged[target].where(merged[target].notna() & merged[target].astype(str).ne(''),merged[c])
            merged=merged.drop(columns=c)
    remaining=remaining.copy();remaining.insert(1,'Тип блока','Дневная детализация')
    out=pd.concat([merged,remaining],ignore_index=True)
    return out.sort_values(['Период',keys[-1]],key=lambda x:x.map(lambda v:wb_period(v)[0].isoformat() if wb_period(v) else str(v)) if x.name=='Период' else x,kind='stable').reset_index(drop=True)


def additive(c):
    # Белый список типов показателей для итогов: проценты, цены, сроки и
    # средние не складываются. Исходные альтернативы всегда раздельны.
    low=c.lower()
    if '|' not in c or any(x in low for x in ('предыдущ','даты не указаны','доля','процент','конверси','средн',' / ','/шт','в день','дней','дня','дн','цена','sku','артикул','рейтинг','индекс','уровень','примечание','источник','дата')):
        return c.endswith(' | Прибыль за период')
    return any(x in low for x in ('шт','₽','руб','rub','выручка','количество','отменено','доставлено','остальные статусы','стоимость','комисси','эквайринг','логистика','баллы','программ','вознаграждение','прибыль','итог','удержани','переходы','корзин'))


def totals(df):
    if df.empty or 'Период' not in df:return pd.DataFrame()
    keys=['Период']+(['Тип блока'] if 'Тип блока' in df else [])+(['Схема'] if 'Схема' in df else [])
    cols=[c for c in df if additive(c) and pd.to_numeric(df[c],errors='coerce').notna().any()]
    if not cols:return pd.DataFrame()
    out=df.groupby(keys,dropna=False,sort=False)[cols].agg(safe_sum).reset_index()
    # Взвешенные доли считаем заново из сумм, не складываем проценты товаров.
    for label,a,b in [('Доставлено / заказано, %','Заказы | Доставлено','Заказы | Количество'),('Отменено / заказано, %','Заказы | Отменено','Заказы | Количество'),('Выкуплено / заказано — WB, %','Воронка | Выкупили, шт','Воронка | Заказали товаров, шт'),('Выкуплено / заказано — поставщик, %','Поставщик | Выкупили, шт.','Поставщик | шт.')]:
        if a in out and b in out:out['Расчёт | '+label]=out[a]/out[b].replace(0,float('nan'))*100
    delivered='Юнит | Доставлено товаров, шт'
    if delivered in out:
        denominator=out[delivered].where(out[delivered]>0)
        if 'Юнит | Выручка' in out:out['Расчёт | Выручка на доставленный товар']=out['Юнит | Выручка']/denominator
        if 'Юнит | Прибыль за период' in out:out['Расчёт | Прибыль на доставленную единицу, ₽/шт']=out['Юнит | Прибыль за период']/denominator
    return out


def analyze_all(paths,progress=lambda text:None):
    paths=list(dict.fromkeys(str(Path(p).resolve()) for p in paths))
    warnings=[];registry=[];unknown={};routes=defaultdict(list);seen={}
    with TemporaryDirectory(prefix='glasok_') as temp:
        expanded=[]
        for i,path in enumerate(paths):
            p=Path(path)
            if p.suffix.lower()=='.zip':
                try:
                    with ZipFile(p) as z:
                        members=[m for m in z.infolist() if Path(m.filename).suffix.lower() in SUPPORTED-{'.zip'} and not m.is_dir() and not Path(m.filename).name.startswith('.')]
                        if sum(m.file_size for m in members)>300_000_000:raise ValueError('Распакованный архив больше 300 МБ')
                        for j,m in enumerate(members):
                            name=m.filename
                            # Some mail services store UTF-8 names as legacy CP437 ZIP names.
                            if not m.flag_bits & 0x800:
                                try:name=m.filename.encode('cp437').decode('utf-8')
                                except (UnicodeError,UnicodeEncodeError):pass
                            name=Path(name).name
                            dest=Path(temp)/str(i)/str(j)/name;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(z.read(m));expanded.append((dest,p.name+' / '+name))
                        if not members:raise ValueError('В архиве нет поддерживаемых таблиц')
                except Exception as e:registry.append({'Файл':p.name,'Маркетплейс':'—','Статус':'Ошибка','Причина':str(e)});warnings.append(f'{p.name}: {e}')
            else:expanded.append((p,p.name))
        for i,(p,label) in enumerate(expanded):
            progress(f'Определение файлов: {i+1}/{len(expanded)} — {label}')
            try:
                if p.suffix.lower() not in SUPPORTED:raise ValueError('Поддерживаются XLSX, XLSM, XLS, CSV, TSV и ZIP')
                digest=hashlib.sha256(p.read_bytes()).hexdigest()
                if digest in seen:
                    registry.append({'Файл':label,'Маркетплейс':seen[digest],'Статус':'Повтор: не суммируется','Причина':'Полное совпадение содержимого'});continue
                market,preview=classify(p);seen[digest]=market
                if market in ('WB','Ozon'):
                    routes[market].append(str(p));registry.append({'Файл':label,'Маркетплейс':market,'Статус':'Определён','Причина':'По заголовкам внутри файла'})
                else:
                    registry.append({'Файл':label,'Маркетплейс':market,'Статус':'Нужен новый обработчик','Причина':'Данные сохранены для просмотра без расчётов'})
                    for sheet,rows in preview.items():unknown[f'{label} / {sheet}']=pd.DataFrame(rows)
                    warnings.append(f'{label}: нет знакомых заголовков; файл доступен в разделе «Не распознано»')
            except Exception as e:
                registry.append({'Файл':label,'Маркетплейс':'—','Статус':'Ошибка','Причина':str(e)});warnings.append(f'{label}: {e}')
        result={}
        for market in ('Ozon','WB'):
            if not routes[market]:result[market]={};continue
            progress(f'{market}: объединение и расчёты…')
            try:
                views,notes=analyze_ozon(routes[market]) if market=='Ozon' else analyze_wb(routes[market],keep_zero=True)
                warnings.extend(notes)
                if market=='WB':
                    full=enrich_wb(views.get('Товары по периодам',pd.DataFrame()));daily=views.get('Пересечения дат',pd.DataFrame())
                    views['Товары по периодам']=full
                else:full=views.get('Сводка полных периодов',pd.DataFrame());daily=views.get('Заказы и возвраты по пересечениям дат',pd.DataFrame())
                combined=join_period_views(full,daily,market)
                result[market]={'Общая сводка товаров':combined,'Итоги по периодам':totals(combined),**views}
                if market=='WB' and not views.get('Финансы магазина',pd.DataFrame()).empty:
                    finance=views['Финансы магазина'];keys=['Период','Финансы | Валюта']
                    exclude=('Дата','№','Юридическое','Тип','Валюта','скидка, %')
                    columns=[c for c in finance if c.startswith('Финансы | ') and not any(x in c for x in exclude) and pd.to_numeric(finance[c],errors='coerce').notna().any()]
                    report=finance.groupby(keys,dropna=False)[columns].agg(safe_sum).reset_index()
                    if 'Финансы | Итого к оплате' in report and 'Финансы | Продажа' in report:
                        report['Расчёт | К оплате / продажи, %']=report['Финансы | Итого к оплате']/report['Финансы | Продажа'].replace(0,float('nan'))*100
                    result[market]['Итоги финансов магазина']=report
                if market=='Ozon':
                    snapshots=snapshot_views(views)
                    if not snapshots.empty:
                        joined=sales_and_stocks(full,snapshots)
                        result[market]={'Общая сводка товаров':combined,
                                        'Продажи и текущие остатки':joined,
                                        'Остатки + маркировка + план поставки':snapshots,**result[market]}
                result[market]={k:v for k,v in result[market].items() if not v.empty}
                for r in registry:
                    if r['Маркетплейс']==market and r['Статус']=='Определён':r['Статус']='Обработан'
            except Exception as e:
                warnings.append(f'{market}: {e}');result[market]={}
                for r in registry:
                    if r['Маркетплейс']==market and r['Статус']=='Определён':r['Статус']='Ошибка обработки';r['Причина']=str(e)
        result['Загрузка']={'Файлы':pd.DataFrame(registry)}
        if warnings:result['Загрузка']['Замечания']=pd.DataFrame({'Сообщение':list(dict.fromkeys(warnings))})
        if unknown:result['Не распознано']=unknown
        return result,warnings


def snapshot_views(views):
    """Остатки и маркировка одной даты, без смешивания старого неизвестного среза."""
    frames=[v.copy() for k,v in views.items() if k.startswith(('Остатки:','Маркировка:'))]
    grouped=defaultdict(list)
    for frame in frames:
        for period,part in frame.groupby('Период',sort=False):grouped[period].append(part)
    output=[]
    for period,parts in grouped.items():
        rows={}
        for part in parts:
            for r in part.to_dict('records'):
                sku=ident(r.get('SKU')) if pd.notna(r.get('SKU')) else '';key=sku or 'article:'+ident(r.get('Артикул'))
                out=rows.setdefault(key,{'Период':period,'SKU':sku,'Артикул':r.get('Артикул',''),'Название':r.get('Название','')})
                for c,v in r.items():
                    if v is None or isinstance(v,float) and math.isnan(v):continue
                    if c in ('Источник','Строка Excel'):
                        c=('Маркировка' if any(k.startswith('Маркировка |') for k in r) else 'Остатки')+' | '+c
                    if c in out and out[c] not in ('',None) and out[c]!=v:
                        if c=='Артикул':out[c]='; '.join(sorted(set(str(out[c]).split('; '))|{str(v)}))
                        elif c not in ('Период','SKU','Название'):out['Конфликт | '+c]=str(v)
                    else:out[c]=v
        for out in rows.values():
            rate=next((num(v) for c,v in out.items() if c.startswith('Остатки | Среднесуточные продажи')),None)
            stock=num(out.get('Остатки | Доступно к продаже'))
            transit=num(out.get('Остатки | В поставках в пути'))
            if rate is not None and rate>0 and stock is not None:
                out['Расчёт | Целевой запас на 28 дней, шт']=math.ceil(rate*28)
                out['Расчёт | Потребность на 28 дней без учёта пути, шт']=max(0,math.ceil(rate*28-stock))
                if transit is not None:out['Расчёт | Потребность на 28 дней с учётом пути, шт']=max(0,math.ceil(rate*28-stock-transit))
                out['Расчёт | Спрос на 30 дней, шт']=math.ceil(rate*30)
                out['Расчёт | Довезти на 30 дней без учёта пути, шт']=max(0,math.ceil(rate*30-stock))
                if transit is not None:
                    out['Расчёт | Довезти на 30 дней с учётом пути, шт']=max(0,math.ceil(rate*30-stock-transit))
                out['Расчёт | Приоритет пополнения']='Менее 7 дней' if stock/rate<7 else 'Менее 28 дней' if stock/rate<28 else 'Запас от 28 дней'
            elif rate==0 and stock is not None and stock>0:out['Расчёт | Приоритет пополнения']='Есть остаток, нет среднесуточных продаж'
            output.append(out)
    return pd.DataFrame(output)


def sales_and_stocks(periods,snapshots):
    """Align a sales period with the latest stock snapshot by SKU; retain both dates."""
    if periods.empty or snapshots.empty or 'SKU' not in periods or 'SKU' not in snapshots:return pd.DataFrame()
    sales=periods.copy()
    stock=snapshots.copy()
    stock['Дата среза остатков']=stock.pop('Период')
    stock=stock[stock['SKU'].fillna('').astype(str).ne('')]
    # Snapshot labels are YYYY-independent strings here; extract ISO dates to sort.
    def stamp(label):
        from marketplace_wb import wb_date
        import re
        match=re.search(r'\d{2}\.\d{2}\.\d{4}',str(label))
        day=wb_date(match.group()) if match else None
        return day.isoformat() if day else ''
    stock['_snapshot_sort']=stock['Дата среза остатков'].map(stamp)
    stock=stock.sort_values('_snapshot_sort',kind='stable').drop_duplicates('SKU',keep='last').drop(columns='_snapshot_sort')
    sales['SKU']=sales['SKU'].fillna('').astype(str)
    stock['SKU']=stock['SKU'].fillna('').astype(str)
    # Keep every period and sales scheme; this is a product view, not a sum
    # across snapshots or a claim that the stock existed on the sales end date.
    result=sales.merge(stock,on='SKU',how='outer',validate='many_to_one',suffixes=('','__stock'))
    for c in ('Артикул','Название'):
        other=c+'__stock'
        if other in result:
            result[c]=result[c].where(result[c].notna() & result[c].astype(str).ne(''),result[other])
            result=result.drop(columns=other)
    # Physical Ozon stock belongs to FBO. Do not attach it to FBS rows.
    if 'Схема' in result:
        foreign=result['Схема'].fillna('').str.upper().isin(['FBS','RFBS'])
        result.loc[foreign,[c for c in stock if c not in ('SKU','Артикул','Название') and c in result]]=None
    rate=pd.to_numeric(result.get('Расчёт | Продажи в день по доставленным, шт',pd.Series(index=result.index,dtype=float)),errors='coerce')
    available=pd.to_numeric(result.get('Остатки | Доступно к продаже',pd.Series(index=result.index,dtype=float)),errors='coerce')
    result['Расчёт | Остаток в днях по доставленным']=available/rate.where(rate>0)
    # A stock-only SKU remains visible even when it had no sales report row.
    return result
