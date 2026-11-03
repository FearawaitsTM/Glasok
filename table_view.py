"""Display-only ordering and filtering; report calculations keep original fields."""
import pandas as pd

from excel_grid import display


HIDDEN = {'Источник набора', 'Тип блока', 'Схема', 'Источник', 'Источники',
          'Доступность | Дата среза', 'Доступность | Источник',
          'Покрытие отчётами', 'Примечание'}
IDENTITY = (
    ('Артикул', ('Артикул', 'Артикул продавца', 'Артикул WB')),
    ('SKU', ('SKU', 'SKU / штрихкод', 'ID товара')),
    ('Название товара', ('Название товара', 'Название', 'Наименование', 'Товар')),
)


def identity_columns(df):
    return [(label, tuple(name for name in aliases if name in df.columns))
            for label, aliases in IDENTITY]


def product_matches(df, query):
    """Case-insensitive substring search in product names and identifiers."""
    names = [name for _, aliases in identity_columns(df) for name in aliases]
    if not names:
        names = list(df.columns)  # File registry and other non-product views.
    if not query or not names:
        return pd.Series(True, index=df.index)
    q = query.casefold()
    return df[names].fillna('').astype(str).apply(
        lambda column: column.str.casefold().str.contains(q, regex=False)
    ).any(axis=1)


def product_suggestions(df, query, limit=5):
    if df.empty:
        return []
    identities = identity_columns(df)
    if not any(aliases for _, aliases in identities):
        return []
    result = []
    seen = set()
    for _, row in df.loc[product_matches(df, query)].iterrows():
        fields = []
        for _, aliases in identities:
            fields.append(next((display(row[name]) for name in aliases
                                if display(row[name]).strip()), ''))
        key = tuple(fields)
        identity = ('article',fields[0]) if fields[0] else ('sku',fields[1]) if fields[1] else ('name',fields[2])
        if key == ('', '', '') or identity in seen:
            continue
        seen.add(identity)
        label = ' · '.join(part for part in (fields[2], f'Арт. {fields[0]}' if fields[0] else '',
                                           f'SKU {fields[1]}' if fields[1] else '') if part)
        result.append((label, key))
        if len(result) == limit:
            break
    return result


def exact_product(df, key):
    article, sku, name = key
    # Match either identifier so rows from a source with only SKU are included.
    matches = []
    for value, aliases in ((article, IDENTITY[0][1]), (sku, IDENTITY[1][1])):
        names = [n for n in aliases if n in df.columns]
        if value and names:
            matches.append(df[names].apply(lambda column:column.map(display).eq(value)).any(axis=1))
    if matches:
        result=matches[0]
        for match in matches[1:]:result=result|match
        return result
    names=[n for n in IDENTITY[2][1] if n in df.columns]
    if name and names:return df[names].apply(lambda column:column.map(display).eq(name)).any(axis=1)
    return pd.Series(False,index=df.index)


def present(df, show_calculations=True):
    """Put product identifiers first, original fields next, derived fields last."""
    df = df.drop(columns=[c for c in HIDDEN if c in df.columns], errors='ignore').copy()
    primary = []
    renames = {}
    for label, aliases in IDENTITY:
        for name in aliases:
            if name in df.columns and name not in primary:
                primary.append(name)
                if label not in df.columns and label not in renames.values():
                    renames[name] = label
                break
    calculations = [c for c in df.columns if str(c).startswith('Расчёт | ')]
    original = [c for c in df.columns if c not in primary and c not in calculations]
    columns = primary + original + (calculations if show_calculations else [])
    result = df.loc[:, columns].rename(columns=renames)
    if primary:
        for label, _ in IDENTITY:
            if label not in result.columns:
                result[label] = ''
        labels = [label for label, _ in IDENTITY]
        result = result.loc[:, labels + [c for c in result if c not in labels]]
    if any(c.startswith('Юнит | ') for c in result):
        labels={
            'Юнит | Передано в доставку покупателям, шт':'Передано в доставку, шт',
            'Юнит | Доставлено товаров, шт':'Доставлено, шт',
            'Юнит | Возвращено товаров, шт':'Возвращено за период, шт',
            'Расчёт | Передано за вычетом возвратов за период, шт':'Расчёт | Итог после возвратов, шт',
            'Остатки | Доступно к продаже':'Остаток, шт',
            'Расчёт | Продажи в день по доставленным, шт':'Расчёт | Продажи в день, шт',
            'Расчёт | Остаток в днях по доставленным':'Расчёт | Остаток в днях',
            'Юнит | Выручка':'Выручка от доставленных, ₽',
            'Расчёт | Выручка на доставленный товар':'Расчёт | Средняя цена продажи, ₽/шт',
            'Юнит | Прибыль за период':'Прибыль из отчёта за период, ₽',
            'Расчёт | Прибыль на доставленную единицу, ₽/шт':'Расчёт | Прибыль на единицу, ₽/шт',
        }
        result=result.rename(columns=labels)
        front=['Артикул','SKU','Название товара','Передано в доставку, шт','Доставлено, шт',
               'Возвращено за период, шт','Расчёт | Итог после возвратов, шт','Остаток, шт',
               'Расчёт | Продажи в день, шт','Расчёт | Остаток в днях',
               'Выручка от доставленных, ₽','Расчёт | Средняя цена продажи, ₽/шт',
               'Прибыль из отчёта за период, ₽','Расчёт | Прибыль на единицу, ₽/шт']
        chosen=[c for c in front if c in result]
        result=result[chosen+[c for c in result if c not in chosen]]
    return result


def apply_column_filters(df, filters):
    for col, values in filters.items():
        if col in df.columns and values is not None:
            df = df[df[col].map(display).isin(values)]
    return df
