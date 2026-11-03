"""HTTP clients and normalisation for Ozon Seller API and WB API."""
from __future__ import annotations

import json
import math
import ssl
import sys
import time
from datetime import date, datetime, timedelta, timezone
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


REPORTS = {
    "Ozon": {
        "products": "Товары и маркировка",
        "stocks": "Остатки",
        "orders": "Заказы FBO + FBS",
        "supplies": "Поставки",
        "finance": "Финансовые операции",
        "returns": "Возвраты",
    },
    "WB": {
        "products": "Карточки и маркировка",
        "stocks": "Остатки",
        "orders": "Заказы",
        "supplies": "Поставки FBS",
        "finance": "Финансовый отчёт",
        "funnel": "Воронка продаж",
    },
}


class ApiProblem(RuntimeError):
    pass


def _request(method, url, headers, body=None, timeout=90):
    data=None if body is None else json.dumps(body, ensure_ascii=False).encode("utf-8")
    req=Request(url, data=data, method=method, headers={"Accept":"application/json", **headers})
    if data is not None:req.add_header("Content-Type", "application/json")
    try:
        # Python's OpenSSL defaults on macOS can differ from the certificates
        # trusted by the user's Keychain (for example on managed networks).
        # Keep full verification, including hostname checks, through the OS.
        context=None
        if sys.platform in ('darwin','win32'):
            import truststore
            context=truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with urlopen(req, timeout=timeout, context=context) as response:
            raw=response.read()
            return json.loads(raw.decode("utf-8")) if raw else {}
    except HTTPError as exc:
        detail=exc.read().decode("utf-8", "replace")[:1200]
        if exc.code==401 and "api-seller.ozon.ru" in url:
            raise ApiProblem("Ozon отклонил Api-Key (HTTP 401). Проверьте, что Client-Id и действующий ключ Seller API взяты из одного кабинета. Введите оба значения заново в «Глазок Сервер» и нажмите «Сохранить».") from exc
        if exc.code==403 and "wildberries.ru" in url and "additional requirements" in detail.lower():
            if "/stocks-report/" in url:
                raise ApiProblem("WB не разрешил загрузить остатки (HTTP 403). Для этого метода нужен Персональный или Сервисный токен WB с категорией «Аналитика». Проверьте права токена в кабинете WB API и сохраните подходящий токен в «Глазок Сервер».") from exc
            raise ApiProblem("WB не разрешил этот отчёт (HTTP 403). Проверьте категорию, тип токена и условия доступа к отчёту в кабинете WB API; затем сохраните подходящий токен в «Глазок Сервер».") from exc
        raise ApiProblem(f"HTTP {exc.code}: {detail or exc.reason}") from exc
    except URLError as exc:
        if 'CERTIFICATE_VERIFY_FAILED' in str(exc.reason) or 'certificate verify failed' in str(exc.reason).lower():
            raise ApiProblem('Не удалось проверить сертификат HTTPS. Проверьте доверие к сертификатам в macOS и настройки VPN или прокси; затем повторите запрос.') from exc
        raise ApiProblem(f"Нет соединения: {exc.reason}") from exc


def _post(url, headers, body):return _request("POST",url,headers,body)
def _get(url, headers, params=None):return _request("GET",url+("?"+urlencode(params) if params else ""),headers)


def _items(payload):
    """Find the most useful list in a changing API response without losing raw fields."""
    if isinstance(payload,list):return payload
    if not isinstance(payload,dict):return [{"value":payload}]
    preferred=("items","result","operations","postings","returns","orders","supplies","cards","products","data","report")
    for key in preferred:
        value=payload.get(key)
        if isinstance(value,list):return value
        if isinstance(value,dict):
            found=_items(value)
            if found:return found
    return [payload] if payload else []


def _flatten(value, prefix=""):
    out={}
    if isinstance(value,dict):
        for key,item in value.items():
            name=f"{prefix}.{key}" if prefix else str(key)
            if isinstance(item,dict):out.update(_flatten(item,name))
            elif isinstance(item,list):
                if item and all(isinstance(x,dict) for x in item):out[name]=json.dumps(item,ensure_ascii=False,separators=(",",":"))
                else:out[name]="; ".join(str(x) for x in item)
            else:out[name]=item
    else:out[prefix or "value"]=value
    return out


def _expand(row):
    """Explode product/size/stock arrays so one result row can be joined by product."""
    if not isinstance(row,dict):return [_flatten(row)]
    priority=("products","items","stocks","sizes")
    candidate=next(((key,row[key]) for key in priority if isinstance(row.get(key),list) and row[key] and all(isinstance(x,dict) for x in row[key])),None)
    if not candidate:return [_flatten(row)]
    key,children=candidate;parent={k:v for k,v in row.items() if k!=key};base=_flatten(parent)
    output=[]
    for child in children:
        nested=_expand(child)
        for detail in nested:output.append({**base,**{f"{key}.{k}":v for k,v in detail.items()}})
    return output


def _flat_records(rows):return [flat for row in rows for flat in _expand(row)]
def _clean_rows(payload):return _flat_records(_items(payload))


def _period(start,end):return f"{start} — {end}"


def fetch_ozon(report, settings, start, end):
    client=str(settings.get("ozon_client_id","")).strip();token=str(settings.get("ozon_api_key","")).strip()
    if not client or not token:raise ApiProblem("В настройках не указаны Client-Id и Api-Key Ozon")
    headers={"Client-Id":client,"Api-Key":token};base="https://api-seller.ozon.ru"
    iso_from=f"{start}T00:00:00.000Z";iso_to=f"{end}T23:59:59.999Z"
    if report=="products":
        rows=[];last_id=""
        for _ in range(100):
            data=_post(base+"/v3/product/list",headers,{"filter":{"visibility":"ALL"},"last_id":last_id,"limit":1000})
            batch=_items(data);rows.extend(batch)
            nxt=((data.get("result") or {}).get("last_id") if isinstance(data,dict) else "") or ""
            if not batch or not nxt or nxt==last_id:break
            last_id=nxt
        return _flat_records(rows)
    if report=="stocks":
        rows=[];cursor=""
        for _ in range(100):
            data=_post(base+"/v4/product/info/stocks",headers,{"filter":{"visibility":"ALL"},"limit":1000,"cursor":cursor})
            batch=_items(data);rows.extend(batch);nxt=(data.get("cursor") or "") if isinstance(data,dict) else ""
            if not batch or not nxt or nxt==cursor:break
            cursor=nxt
        return _flat_records(rows)
    if report=="orders":
        rows=[]
        for scheme,path in (("FBO","/v3/posting/fbo/list"),("FBS","/v4/posting/fbs/list")):
            cursor=""
            for _ in range(100):
                body={"sort_dir":"asc","filter":{"since":iso_from,"to":iso_to},"limit":100,"cursor":cursor,"with":{"analytics_data":True,"financial_data":True}}
                data=_post(base+path,headers,body);batch=_items(data)
                rows.extend([{"scheme":scheme,**x} if isinstance(x,dict) else {"scheme":scheme,"value":x} for x in batch])
                if not data.get("has_next",False):break
                nxt=data.get("cursor") or ""
                if not nxt or nxt==cursor:raise ApiProblem(f"Ozon {scheme}: ответ содержит следующую страницу без нового курсора")
                cursor=nxt
            else:raise ApiProblem(f"Ozon {scheme}: больше 100 страниц заказов; сузьте период")
        return _flat_records(rows)
    if report=="finance":
        rows=[];day=date.fromisoformat(start);last_day=date.fromisoformat(end)
        while day<=last_day:
            cursor=""
            for _ in range(100):
                data=_post(base+"/v1/finance/accrual/by-day",headers,{"date":day.isoformat(),"last_id":cursor})
                batch=data.get("accruals",[]) or []
                rows.extend([{"date":day.isoformat(),**x} if isinstance(x,dict) else {"date":day.isoformat(),"value":x} for x in batch])
                nxt=data.get("last_id") or ""
                if not nxt:break
                if nxt==cursor:raise ApiProblem(f"Ozon финансы {day}: курсор не изменился")
                cursor=nxt
            else:raise ApiProblem(f"Ozon финансы {day}: больше 100 страниц; сузьте период")
            day+=timedelta(days=1)
        return _flat_records(rows)
    if report=="supplies":
        rows=[];last_id=""
        for _ in range(100):
            body={"filter":{"states":[]},"limit":100,"last_id":last_id}
            data=_post(base+"/v3/supply-order/list",headers,body);batch=_items(data);rows.extend(batch)
            nxt=((data.get("result") or {}).get("last_id") if isinstance(data,dict) else "") or ""
            if len(batch)<100 or not nxt or nxt==last_id:break
            last_id=nxt
        return _flat_records(rows)
    if report=="returns":
        # Endpoint supports period pagination in current Seller API. If an account has no
        # access, the original API error is shown next to this report, not hidden.
        rows=[];last_id=0
        for _ in range(100):
            body={"filter":{"accepted_from":iso_from,"accepted_to":iso_to},"limit":500,"last_id":last_id}
            data=_post(base+"/v1/returns/list",headers,body);batch=_items(data);rows.extend(batch)
            nxt=(data.get("last_id") or 0) if isinstance(data,dict) else 0
            if len(batch)<500 or not nxt or nxt==last_id:break
            last_id=nxt
        return _flat_records(rows)
    raise ApiProblem("Неизвестный отчёт Ozon")


def fetch_wb(report, settings, start, end):
    token=str(settings.get("wb_api_key","")).strip()
    if not token:raise ApiProblem("В настройках не указан токен WB")
    headers={"Authorization":token};period=_period(start,end)
    if report=="products":
        rows=[];cursor={"limit":100}
        for _ in range(100):
            body={"settings":{"cursor":cursor,"filter":{"withPhoto":-1},"sort":{"ascending":False}}}
            data=_post("https://content-api.wildberries.ru/content/v2/get/cards/list",headers,body);batch=_items(data);rows.extend(batch)
            nxt=((data.get("cursor") or {}) if isinstance(data,dict) else {})
            if len(batch)<100 or not nxt.get("updatedAt") or not nxt.get("nmID"):break
            cursor={"limit":100,"updatedAt":nxt["updatedAt"],"nmID":nxt["nmID"]}
        return _flat_records(rows)
    if report=="orders":
        data=_get("https://statistics-api.wildberries.ru/api/v1/supplier/orders",headers,{"dateFrom":start,"flag":0})
        return _clean_rows(data)
    if report=="supplies":
        rows=[];next_id=0
        for _ in range(100):
            data=_get("https://marketplace-api.wildberries.ru/api/v3/supplies",headers,{"limit":1000,"next":next_id});batch=_items(data);rows.extend(batch)
            nxt=(data.get("next") or 0) if isinstance(data,dict) else 0
            if len(batch)<1000 or not nxt or nxt==next_id:break
            next_id=nxt
        return _flat_records(rows)
    if report=="stocks":
        rows=[];offset=0
        for _ in range(100):
            body={"nmIds":[],"chrtIds":[],"limit":250000,"offset":offset}
            data=_post("https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses",headers,body);batch=_items(data);rows.extend(batch)
            if len(batch)<250000:break
            offset+=len(batch)
        return _flat_records(rows)
    if report=="finance":
        rows=[];cursor={}
        for _ in range(100):
            body={"dateFrom":start,"dateTo":end,"limit":100000,"cursor":cursor}
            data=_post("https://finance-api.wildberries.ru/api/finance/v1/sales-reports/detailed",headers,body);batch=_items(data);rows.extend(batch)
            nxt=(data.get("cursor") or {}) if isinstance(data,dict) else {}
            if not batch or not nxt or nxt==cursor:break
            cursor=nxt
        return [_flatten(x) for x in rows]
    if report=="funnel":
        body={"selectedPeriod":{"start":start,"end":end},"pastPeriod":{"start":start,"end":end},"nmIds":[],"subjectIds":[],"brandNames":[],"tagIds":[],"skipDeletedNm":False,"orderBy":{"field":"orders","mode":"desc"},"limit":1000,"offset":0}
        return _clean_rows(_post("https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products",headers,body))
    raise ApiProblem("Неизвестный отчёт WB")


ALIASES={
    "Артикул продавца":("offer_id","offerId","vendorCode","supplierArticle","article"),
    "ID товара":("product_id","productId","nmID","nmId"),
    "SKU / штрихкод":("sku","barcode","barcodes","skus"),
    "Название":("name","title","product_name","productName","subjectName"),
    "Дата":("date","created_at","createdAt","in_process_at","posting_date","operation_date"),
}


def _first(row,names):
    for exact in names:
        for key,value in row.items():
            leaf=key.rsplit(".",1)[-1]
            if leaf==exact and value not in (None,""):return value
    return ""


def build_views(market, datasets, start, end):
    """Return raw views, a product/date outer merge, numeric totals and errors."""
    views={};wide={};period=_period(start,end)
    for code,rows in datasets.items():
        title=REPORTS[market][code];normal=[]
        for source in rows:
            row={"Период":period,"Отчёт":title}
            for label,aliases in ALIASES.items():row[label]=_first(source,aliases)
            row.update({f"{title} | {key}":value for key,value in source.items()})
            normal.append(row)
            product=str(row["Артикул продавца"] or row["ID товара"] or row["SKU / штрихкод"]).strip()
            day=str(row["Дата"] or "").strip()[:10]
            if product:
                key=(period,product,day)
                target=wide.setdefault(key,{"Период":period,"Ключ товара":product,"Дата":day})
                for base in ("Артикул продавца","ID товара","SKU / штрихкод","Название"):
                    if row.get(base) not in (None,""):target.setdefault(base,row[base])
                for col,val in row.items():
                    if " | " in col and val not in (None,""):
                        if col in target and target[col]!=val:
                            current=target[col]
                            if not isinstance(current,list):target[col]=[current]
                            if val not in target[col]:target[col].append(val)
                        else:target[col]=val
        views[f"API — {title}"]=normal
    combined=list(wide.values())
    # Practical derived values are only emitted when their inputs are unambiguous.
    for row in combined:
        numeric={k:_number(v) for k,v in row.items()}
        stock=_pick_number(numeric,("present","available","stock","quantity"),("Остатки",))
        ordered=_pick_number(numeric,("quantity","ordersCount","orders"),("Заказы","Воронка продаж"))
        revenue=_pick_number(numeric,("amount","revenue","price"),("Финансовые операции","Финансовый отчёт","Воронка продаж"))
        if ordered is not None and ordered>0 and revenue is not None:row["Расчёт | Средняя сумма на единицу"]=revenue/ordered
        if stock is not None and ordered is not None:
            days=max(1,(date.fromisoformat(end)-date.fromisoformat(start)).days+1);rate=ordered/days
            if rate>0:
                row["Расчёт | Среднее в день"]=rate;row["Расчёт | Запас, дней"]=stock/rate
                row["Расчёт | Потребность на 28 дней"]=max(0,math.ceil(rate*28-stock))
    views={"API — Объединённая сводка":combined,**views}
    all_rows=[row for rows in datasets.values() for row in rows]
    totals={"Период":period,"Отчётов":len(datasets),"Исходных строк":len(all_rows),"Товаров в сводке":len(combined)}
    for code,rows in datasets.items():totals[f"Строки | {REPORTS[market][code]}"]=len(rows)
    views["API — Итоги"]=[totals]
    return views


def _number(value):
    if isinstance(value,bool) or isinstance(value,(dict,list)):return None
    if isinstance(value,(int,float)):return float(value) if math.isfinite(float(value)) else None
    try:return float(str(value).replace(" ","").replace(",","."))
    except (TypeError,ValueError):return None


def _pick_number(values,leaves,prefixes):
    for key,value in values.items():
        if value is None or not any(prefix in key for prefix in prefixes):continue
        leaf=key.rsplit(".",1)[-1].lower()
        if any(name.lower()==leaf for name in leaves):return value
    return None


def fetch_selected(market,reports,settings,start,end,progress=lambda text:None):
    datasets={};errors=[]
    fetcher=fetch_ozon if market=="Ozon" else fetch_wb
    for index,report in enumerate(reports,1):
        if report not in REPORTS.get(market,{}):errors.append(f"{report}: неизвестный отчёт");continue
        progress(f"{market}: {REPORTS[market][report]} ({index}/{len(reports)})")
        try:datasets[report]=fetcher(report,settings,start,end)
        except Exception as exc:
            errors.append(f"{REPORTS[market][report]}: {exc}")
            if market=="Ozon" and "Ozon отклонил Api-Key" in str(exc):break
        if index<len(reports):time.sleep(.25)
    return build_views(market,datasets,start,end),errors
