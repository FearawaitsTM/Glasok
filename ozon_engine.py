"""Чтение Ozon без смешивания периодов, валют и уровней детализации."""
import csv
import hashlib
import re
from collections import defaultdict
from datetime import timedelta
from pathlib import Path
import pandas as pd
from marketplace_wb import load_workbook_safely, wb_date, wb_period, wb_label


def clean(v):
    if v is None:return ''
    return re.sub(r'\s+', ' ', str(v).replace('_x000D_', ' ')).strip()


def ident(v):
    if isinstance(v,float) and v.is_integer():return str(int(v))
    return clean(v)


def num(v):
    if isinstance(v,(float,int)):return float(v)
    t=clean(v).replace(' ', '').replace(',', '.').replace('руб.', '').replace('₽','')
    try:return float(t)
    except (ValueError,TypeError):return None


def read_reports(paths):
    reports=[];warnings=[];fingerprints=set()
    for path in paths:
        p=Path(path)
        try:
            digest=hashlib.sha256(p.read_bytes()).hexdigest()
            if digest in fingerprints:
                warnings.append(f'{p.name}: повторный файл пропущен');continue
            fingerprints.add(digest)
            w=load_workbook_safely(p)
            try:
                sheets=[]
                for sheet in w:
                    if hasattr(sheet,'reset_dimensions'):sheet.reset_dimensions()
                    sheets.append((sheet.title,list(sheet.values)))
            finally:w.close()
            accepted=0
            for title,rows in sheets:
                h=next((i for i,r in enumerate(rows[:15]) if 'SKU' in r or ('Артикул' in r and 'Товар' in r)),None)
                if h is None:continue
                headers=[clean(v) for v in rows[h]]
                meta=' '.join(clean(v) for r in rows[:h] for v in r if v)
                period=next((wb_period(' '.join(clean(v) for v in row)) for row in rows[:h] if 'период' in ' '.join(clean(v) for v in row).lower() and wb_period(' '.join(clean(v) for v in row))),None)
                if 'Номер заказа' in headers:kind='Заказы';start=h+1
                elif 'Дата возврата' in headers:kind='Возвраты';start=h+1
                elif 'Номер поставки' in headers:kind='Поставки';start=h+1
                elif any('Уровень выкупа товара' in c for c in headers):kind='Выкупы';start=h+2
                elif 'Себестоимость' in headers:kind='Юнит';start=h+1
                elif 'Статус проверки' in headers:kind='Маркировка';start=h+2
                elif any('Рекомендуемая поставка' in c for c in headers):kind='Доступность';start=h+1
                elif title in ('Товары','Товар-склад','Товар-кластер','Товар-склад снято с продажи'):
                    kind='Снято с продажи' if title=='Товар-склад снято с продажи' else 'Остатки';start=h+4
                    headers=[clean(b) or a for a,b in zip(headers,rows[h+1])]
                else:
                    warnings.append(f'{p.name} / {title}: неизвестный тип таблицы, расчёты не выполнены')
                    kind='Прочее';start=h+1
                # Не теряем пустые/повторяющиеся заголовки.
                seen=defaultdict(int);unique=[]
                for j,c in enumerate(headers):
                    c=c or f'Колонка {j+1}';seen[c]+=1
                    unique.append(c if seen[c]==1 else f'{c} [{seen[c]}]')
                records=[]
                for rownum,r in enumerate(rows[start:],start+1):
                    d=dict(zip(unique,r))
                    key=ident(d.get('SKU',d.get('Артикул')))
                    if not key or key in ('—','-') or not re.fullmatch(r'[\w.-]+',key):continue
                    for c in ('SKU','Артикул','Артикул товара','Номер отправления','Номер заказа'):
                        if c in d:d[c]=ident(d[c])
                    d['Файл']=p.name;d['Лист']=title;d['Строка Excel']=rownum
                    records.append(d)
                if not records:continue
                dates=re.findall(r'\d{2}\.\d{2}\.\d{4}',meta+' '+p.name)
                snap=wb_date(dates[-1]) if dates else None
                if kind in ('Заказы','Возвраты'):
                    datecol='Принят в обработку' if kind=='Заказы' else 'Дата возврата'
                    actual=[wb_date(d.get(datecol)) for d in records];actual=[d for d in actual if d]
                    period=(min(actual),max(actual)) if actual else None
                report={'kind':kind,'sheet':title,'rows':records,'period':period,'snapshot':snap,'file':p.name}
                reports.append(report);accepted+=1
            if not accepted:warnings.append(f'{p.name}: формат Ozon не распознан. Отчёты WB загружайте во вкладке Wildberries.')
        except Exception as exc:warnings.append(f'{p.name}: {exc}')
    return reports,warnings


def visible(df,hide=True):
    if not hide or df.empty:return df
    keep=[]
    for c in df:
        s=df[c].dropna();s=s[s.astype(str).str.strip().ne('')]
        if s.empty:continue
        n=pd.to_numeric(s,errors='coerce')
        if n.notna().all() and n.eq(0).all() and c not in ('SKU','Артикул'):continue
        keep.append(c)
    return df[keep]


def analyze_ozon(paths):
    reports,warnings=read_reports(paths)
    views={};catalog=defaultdict(set)
    for rp in reports:
        for r in rp['rows']:
            sku=ident(r.get('SKU'));article=ident(r.get('Артикул',r.get('Артикул товара')))
            if sku and article:catalog[sku].add(article)
    def product(r):
        sku=ident(r.get('SKU'));a=ident(r.get('Артикул',r.get('Артикул товара')))
        if not a and len(catalog[sku])==1:a=next(iter(catalog[sku]))
        return ('Артикул:'+a if a else 'SKU:'+sku),a,sku
    inventory=[];stock_totals=[];full={};daily={};raw=defaultdict(list);seen_reports=set();event_seen={};full_seen={}
    latest_availability={}
    for rp in reports:
        if rp['kind']=='Доступность':
            k=rp['period'];stamp=str(rp['snapshot'] or '')
            latest_availability[k]=max(stamp,latest_availability.get(k,''))
    boundaries=sorted({d for rp in reports if rp['period'] for d in (rp['period'][0],rp['period'][1]+timedelta(days=1))})
    def block(d):
        if not d:return 'Дата не определена'
        for a,b in zip(boundaries,boundaries[1:]):
            if a<=d<b:return wb_label((a,b-timedelta(days=1)))
        return d.isoformat()
    def base(r,label):
        key,a,sku=product(r)
        return {'Период':label,'Артикул':a,'SKU':sku,'Название':clean(r.get('Название товара',r.get('Наименование товара',r.get('Товар',''))))}
    def put(target,r,label,scheme=''):
        key,a,sku=product(r);k=(label,key,scheme)
        out=target.setdefault(k,base(r,label))
        if scheme:out['Схема']=scheme
        if sku:
            skus=set(filter(None,out['SKU'].split(', ')));skus.add(sku);out['SKU']=', '.join(sorted(skus))
        if not out['Название']:out['Название']=base(r,label)['Название']
        return out
    def add(out,col,v):
        n=num(v)
        if n is not None:out[col]=out.get(col,0)+n
    for rp in reports:
        kind=rp['kind'];period=rp['period'];label=wb_label(period) if period else (f"Срез {rp['snapshot']:%d.%m.%Y}" if rp['snapshot'] else 'Дата не указана: '+rp['file'])
        # Содержательные копии файлов с разным именем также не удваиваются.
        signature=(kind,rp['sheet'],label,repr([{k:v for k,v in r.items() if k!='Файл'} for r in rp['rows']]))
        if signature in seen_reports:
            warnings.append(f"{rp['file']} / {rp['sheet']}: совпадающие данные пропущены");continue
        seen_reports.add(signature)
        raw[kind+' — '+rp['sheet']].extend([{'Период':label,**r} for r in rp['rows']])
        inventory.append({'Файл':rp['file'],'Тип':kind,'Лист':rp['sheet'],'Период':label,'Строк':len(rp['rows'])})
        if kind=='Остатки' and rp['sheet']!='Товары':continue
        if kind in ('Остатки','Маркировка','Доступность','Поставки','Прочее','Снято с продажи'):
            # Срезы не выдаём за оборот за период. Разные файлы сохраняем отдельно.
            records=[]
            for r in rp['rows']:
                if kind=='Поставки' and r.get('Номер поставки') not in (None,''):continue
                o=base(r,label)
                o.update({kind+' | '+c:v for c,v in r.items() if c not in ('SKU','Артикул','Название товара','Файл','Лист','Строка Excel')})
                o['Источник']=rp['file'];o['Строка Excel']=r['Строка Excel']
                if kind=='Остатки':
                    stock=num(r.get('Доступно к продаже'));rate=next((num(v) for c,v in r.items() if c.startswith('Среднесуточные продажи')),None)
                    if stock is not None and rate and rate>0:o['Расчёт | Запас, дней']=stock/rate
                    # Эти числа — состояния на дату выгрузки. Они не показывают
                    # историю доставок или причины возврата покупателя.
                    outbound=num(r.get('Доставляем покупателям'))
                    inbound=num(r.get('Возвращаются от покупателей'))
                    if outbound is not None:o['Расчёт | Сейчас в доставке покупателям, шт']=outbound
                    if inbound is not None:o['Расчёт | Сейчас возвращаются от покупателей, шт']=inbound
                    if outbound is not None and inbound is not None:
                        o['Расчёт | Сейчас в доставке и возвращаются, шт']=outbound+inbound
                    total=num(r.get('Всего товаров'))
                    if total is not None and total>0 and stock is not None:
                        o['Расчёт | Доступно от всего учтённого, %']=100*stock/total
                records.append(o)
            views[f'{kind}: {rp["file"]}']=pd.DataFrame(records)
            if kind=='Остатки':
                totals={'Период':label,'Файл':rp['file'],'Товаров (SKU)':len(records)}
                for source,target in (
                    ('Всего товаров','Остатки | Всего товаров, шт'),
                    ('Доступно к продаже','Остатки | Доступно к продаже, шт'),
                    ('Доставляем покупателям','Расчёт | Сейчас в доставке покупателям, шт'),
                    ('Возвращаются от покупателей','Расчёт | Сейчас возвращаются от покупателей, шт'),
                ):
                    values=[num(r.get(source)) for r in rp['rows']]
                    if any(v is not None for v in values):totals[target]=sum(v for v in values if v is not None)
                if all(c in totals for c in ('Расчёт | Сейчас в доставке покупателям, шт','Расчёт | Сейчас возвращаются от покупателей, шт')):
                    totals['Расчёт | Сейчас в доставке и возвращаются, шт']=totals['Расчёт | Сейчас в доставке покупателям, шт']+totals['Расчёт | Сейчас возвращаются от покупателей, шт']
                stock_totals.append(totals)
            if kind=='Доступность' and period and str(rp['snapshot'] or '')==latest_availability.get(period):
                for r in rp['rows']:
                    scheme=clean(r.get('Схема продаж')).upper()
                    key=('availability',period,ident(r.get('SKU')),clean(r.get('Кластер')),scheme)
                    signature=repr({k:v for k,v in r.items() if k not in ('Файл','Лист','Строка Excel')})
                    if key in full_seen:
                        if full_seen[key]!=signature:warnings.append(f'{rp["file"]}: конфликт доступности по SKU/кластеру/схеме; сохранена первая запись')
                        continue
                    full_seen[key]=signature
                    o=put(full,r,label,scheme)
                    for c in ('Рекомендуемая поставка, шт на 28 дней','Среднесуточные продажи, руб. за 28дн','Среднесуточные продажи, шт. за 28дн','Остаток FBO, шт','Остаток FBS, шт','Товары в пути на склад озон, шт'):
                        add(o,'Доступность | '+c,r.get(c))
            continue
        for r in rp['rows']:
            scheme=clean(r.get('Схема продаж',r.get('Схема работы',r.get('Схема','')))).upper()
            if kind in ('Заказы','Возвраты'):
                datecol='Принят в обработку' if kind=='Заказы' else 'Дата возврата'
                d=wb_date(r.get(datecol))
                eventkey=(kind,r.get('Номер отправления'),product(r)[0],r.get('SKU'),r.get('Штрихкод возврата') if kind=='Возвраты' else '')
                sig=repr({k:v for k,v in r.items() if k not in ('Файл','Лист','Строка Excel')})
                if eventkey in event_seen:
                    if event_seen[eventkey]!=sig:warnings.append(f'{rp["file"]}: конфликт записи {eventkey[1]}; в сводке оставлена первая, обе доступны в исходных строках')
                    continue
                event_seen[eventkey]=sig
                # В CSV схема не указана: не угадываем её по складу.
                o=put(daily,r,block(d))
                if kind=='Возвраты':add(o,'Возвраты | Количество',r.get('Количество возвращаемых товаров'))
                else:
                    q=num(r.get('Количество'));add(o,'Заказы | Количество',q)
                    status=clean(r.get('Статус')).lower().replace('ё','е')
                    field='Доставлено' if status=='доставлен' else 'Отменено' if status=='отменен' else 'Остальные статусы'
                    add(o,'Заказы | '+field,q)
                    currency=clean(r.get('Код валюты товара')) or 'валюта не указана'
                    price=num(r.get('Предельная цена'))
                    if q is not None and price is not None:
                        add(o,f'Расчёт | Стоимость заказанных товаров, {currency}',price*q)
                        if field=='Доставлено':add(o,f'Расчёт | Стоимость доставленных товаров, {currency}',price*q)
                continue
            key=(kind,label,product(r)[0],ident(r.get('SKU')),scheme,clean(r.get('Товар','')))
            signature=repr({k:v for k,v in r.items() if k not in ('Файл','Лист','Строка Excel')})
            previous=full_seen.get(key)
            if previous and previous['file']!=rp['file']:
                if signature not in previous['signatures']:warnings.append(f'{rp["file"]}: конфликт {kind} за {label}; первая версия файла остаётся в сводке, все версии — в исходных')
                continue
            if not previous:full_seen[key]={'file':rp['file'],'signatures':set()}
            full_seen[key]['signatures'].add(signature)
            # Один набор товар + схема + полный период. Источники остаются раздельными колонками.
            o=put(full,r,label,scheme)
            source=kind+' | '
            o['Источники']=', '.join(sorted(set(filter(None,o.get('Источники','').split(', ')))|{rp['file']}))
            additive=(['Заказано товаров','Отменено (из Заказанных)','Возвращено (из Заказанных)','Сумма отмен и возвратов'] if kind=='Выкупы' else ['Заказано товаров, шт','Передано в доставку покупателям, шт','Доставлено товаров, шт','Возвращено товаров, шт','Выручка','Баллы за скидки','Программы партнёров','Вознаграждение Ozon','Эквайринг','Обработка отправления','Логистика','Доставка до места выдачи','Стоимость размещения','Обработка возврата','Обратная логистика','Утилизация','Дополнительная обработка ОВХ','Операционные ошибки','Оплата за клик','Оплата за заказ','Звёздные товары','Платный бренд','Отзывы','Прибыль за период'])
            for c in additive:add(o,source+c,r.get(c))
            for c in r:
                if c.startswith('Уровень выкупа товара'):
                    # Это скользящие последние 50 заказов, не процент за период.
                    o[source+c]='; '.join(sorted(set(filter(None,o.get(source+c,'').split('; ')))|{clean(r[c])}))
            if period:o['_days']=(period[1]-period[0]).days+1
    for o in full.values():
        days=o.pop('_days',None)
        for prefix,ordered,cancelled,returned in [('Выкупы','Заказано товаров','Отменено (из Заказанных)','Возвращено (из Заказанных)')]:
            q=o.get(prefix+' | '+ordered);c=o.get(prefix+' | '+cancelled);r=o.get(prefix+' | '+returned)
            if q and q>0:
                if c is not None:o['Расчёт | Доля отмен, %']=100*c/q
                if r is not None:o['Расчёт | Доля возвратов, %']=100*r/q
                if c is not None and r is not None:o['Расчёт | Без отмен и возвратов, шт']=q-c-r
                if days:o['Расчёт | Заказано в день (выкупы)']=q/days
        rev=o.get('Юнит | Выручка');delivered=o.get('Юнит | Доставлено товаров, шт');ordered=o.get('Юнит | Заказано товаров, шт');profit=o.get('Юнит | Прибыль за период')
        if delivered and rev is not None:o['Расчёт | Выручка на доставленный товар']=rev/delivered
        if delivered and delivered>0 and profit is not None:o['Расчёт | Прибыль на доставленную единицу, ₽/шт']=profit/delivered
        if delivered is not None and days:o['Расчёт | Продажи в день по доставленным, шт']=delivered/days
        if rev and profit is not None:o['Расчёт | Прибыль отчёта / выручка, %']=profit/rev*100
        if ordered and delivered is not None:o['Расчёт | Доставлено / заказано, %']=delivered/ordered*100
        if ordered is not None and days:o['Расчёт | Заказано в день (юнит)']=ordered/days
        handed=o.get('Юнит | Передано в доставку покупателям, шт')
        returned=o.get('Юнит | Возвращено товаров, шт')
        if handed is not None and returned is not None:
            o['Расчёт | Передано за вычетом возвратов за период, шт']=handed-returned
        if handed is not None and delivered is not None:
            o['Расчёт | Передано минус доставлено за период, шт']=handed-delivered
            if handed>0:o['Расчёт | Доставлено / передано в доставку, %']=delivered/handed*100
        if returned is not None and delivered and delivered>0:
            o['Расчёт | Возвращено / доставлено по отчёту, %']=returned/delivered*100
    for o in daily.values():
        q=o.get('Заказы | Количество');d=o.get('Заказы | Доставлено',0);c=o.get('Заказы | Отменено',0)
        if q and q>0:
            o['Расчёт | Доставлено / заказано, %']=d/q*100;o['Расчёт | Отменено / заказано, %']=c/q*100
        # Не делим возвраты по дате возврата на заказы по дате оформления: это разные когорты.
    def ordered_frame(records):
        records=list(records)
        def key(r):
            p=wb_period(r['Период']);return (p[0].isoformat() if p else '9999',r['Артикул'],r.get('Схема',''))
        frame=pd.DataFrame(sorted(records,key=key))
        if frame.empty:return frame
        # Put the user's main unit calculation next to its two source numbers.
        front=('Период','Артикул','SKU','Название','Схема',
               'Юнит | Передано в доставку покупателям, шт',
               'Юнит | Возвращено товаров, шт',
               'Расчёт | Передано за вычетом возвратов за период, шт')
        chosen=[c for c in front if c in frame]
        return frame[chosen+[c for c in frame if c not in chosen]]
    result={'Сводка полных периодов':ordered_frame(full.values()),'Заказы и возвраты по пересечениям дат':ordered_frame(daily.values()),'Итоги остатков по срезам':pd.DataFrame(stock_totals),**views,'Загруженные отчёты':pd.DataFrame(inventory)}
    result.update({'Исходные | '+k:pd.DataFrame(v) for k,v in raw.items()})
    return {k:v for k,v in result.items() if not v.empty},warnings
