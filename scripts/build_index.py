# -*- coding: utf-8 -*-
"""
build_index.py — индекс проекта из TSV-выгрузки Лоцман.

Вход  (папка TSV, cp1251):
    objects.tsv     id|kind|parent|type|product|version|state|linktype|name|note
    files.tsv       docId, docType, docProduct, docVer, fileId, fileName, size, crc, modified
    files_done.tsv  то же + status

Выход (папка проекта):
    index.md        дерево объектов с гиперссылками ask:Loodsman.URL
    versions.json   machine-readable состояние для проверки версий
    *.url           гиперссылка объекта в его папке

Раскладка файлов совпадает с loodsman_export.vbs fetch: папка объекта
(Обозначение + Наименование + примечание), вложенность = состав; файлы
документов лежат прямо в папке объекта-владельца.

Запуск:
    python build_index.py [tsvDir] [projectDir] [db] [treeFolder] [rootId] [rootTitle]
Без аргументов значения берутся из local/config.json.
"""
import base64
import json
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(SCRIPT_DIR)
CONFIG = os.path.join(SKILL_ROOT, 'local', 'config.json')

cfg = {}
if os.path.exists(CONFIG):
    with open(CONFIG, encoding='utf-8') as f:
        cfg = json.load(f)


def arg(n, key, default=''):
    if len(sys.argv) > n:
        return sys.argv[n]
    return cfg.get(key, default)


TSV = arg(1, 'tsv_dir')
OUT = arg(2, 'project_dir')
DB = arg(3, 'db')
FOLDER = str(arg(4, 'tree_folder'))
ROOT_ID = str(arg(5, 'root_id'))
ROOT_TITLE = arg(6, 'root_title', ROOT_ID)

if not (TSV and OUT and DB and FOLDER and ROOT_ID):
    sys.exit('Заполните local/config.json (tsv_dir, project_dir, db, tree_folder, root_id) '
             'или передайте аргументы: tsvDir projectDir db treeFolder rootId [rootTitle]')


def read_tsv(path):
    rows = []
    with open(path, encoding='cp1251') as f:
        next(f)
        for line in f:
            rows.append(line.rstrip('\r\n').split('\t'))
    return rows


def make_link(obj_id):
    raw = f'0|{DB}|{FOLDER}|{obj_id}'.encode('cp1251')
    b64 = base64.b64encode(raw).decode('ascii')
    return f'ask:Loodsman.URL?Action=Navigate,params={b64}'


def safe(s):
    for ch in '\\/:*?"<>|\t':
        s = s.replace(ch, '_')
    return s[:120]


# ---- загрузка
objects = {}   # id -> dict
for line in open(os.path.join(TSV, 'objects.tsv'), encoding='cp1251').readlines()[1:]:
    p = line.rstrip('\r\n').split('|')
    if len(p) < 7:
        continue
    oid, kind, parent, otype, product, ver, state, link = (p + [''] * 8)[:8]
    name = p[8] if len(p) > 8 else ''
    note = p[9] if len(p) > 9 else ''
    objects[oid] = dict(id=oid, kind=kind, parent=parent, type=otype,
                        product=product, ver=ver, state=state, link=link,
                        name=name, note=note)

files = []
fetched = {}
for p in read_tsv(os.path.join(TSV, 'files.tsv')):
    if len(p) < 9:
        continue
    doc_id, doc_type, doc_prod, doc_ver, file_id, name, size, crc, modified = p[:9]
    files.append(dict(doc_id=doc_id, doc_type=doc_type, doc_prod=doc_prod,
                      doc_ver=doc_ver, file_id=file_id, name=name,
                      size=size, crc=crc, modified=modified))

done_path = os.path.join(TSV, 'files_done.tsv')
if os.path.exists(done_path):
    for p in read_tsv(done_path):
        if len(p) >= 10:
            fetched[p[5] + '|' + p[0]] = p[9]


# ---- раскладка: папка объекта = Обозначение Наименование Примечание
def display(o):
    r = o['product'] or ''
    for extra in (o['name'], o['note']):
        if extra and extra.lower() not in r.lower():
            r = f'{r} {extra}'.strip()
    if not r or r == 'ROOT':
        r = f'{o["type"]}_{o["id"]}'
    return safe(r)


def rel_dir(oid):
    """Относительный путь папки объекта (root -> '')."""
    chain = []
    cur = oid
    for _ in range(60):
        if cur == ROOT_ID or cur not in objects:
            break
        chain.append(display(objects[cur]))
        cur = objects[cur]['parent']
    return os.path.join(*reversed(chain)) if chain else ''


for fl in files:
    owner = objects.get(fl['doc_id'], {}).get('parent', '')
    fl['rel_dir'] = rel_dir(owner) if owner and owner != ROOT_ID else ''
    fl['rel_path'] = (fl['rel_dir'] + '\\' if fl['rel_dir'] else '') + safe(fl['name'])

# ---- versions.json (для проверки версий)
index = {
    'base': DB, 'tree_folder': FOLDER, 'root': ROOT_ID, 'root_title': ROOT_TITLE,
    'objects': [dict(id=o['id'], kind=o['kind'], parent=o['parent'], type=o['type'],
                     product=o['product'], version=o['ver'], state=o['state'],
                     name=o['name'], note=o['note'],
                     rel_dir=rel_dir(o['id']) if o['id'] != ROOT_ID else '',
                     url=make_link(o['id']))
                for o in objects.values()],
    'files': [dict(f, url=make_link(f['doc_id']),
                   downloaded=fetched.get(f['name'] + '|' + f['doc_id'], '')) for f in files],
}
with open(os.path.join(OUT, 'versions.json'), 'w', encoding='utf-8') as f:
    json.dump(index, f, ensure_ascii=False, indent=1)

# ---- index.md
docs = {}   # doc id -> [files]
for fl in files:
    docs.setdefault(fl['doc_id'], []).append(fl)

lines = [f'# Индекс проекта: {ROOT_TITLE}', '']
lines.append(f'База: {DB}, узел дерева: {FOLDER}. Объектов: {len(objects)}, файлов: {len(files)}.')
lines.append('Файлы лежат в папке объекта-владельца (Обозначение Наименование Примечание).')
lines.append('Ссылки вида `ask:Loodsman.URL?...` открывают объект в ЛОЦМАН Клиент.')
lines.append('')

def render_obj(oid, depth, out):
    o = objects.get(oid)
    if o is None:
        return
    if oid == ROOT_ID:
        o = dict(o, product=ROOT_TITLE)
    pad = '  ' * depth
    state = f' ({o["state"]})' if o['state'] else ''
    dirname = rel_dir(oid) if oid != ROOT_ID else '(корень)'
    title = display(o) if oid != ROOT_ID else ROOT_TITLE
    out.append(f'{pad}- **{title or o["type"]}** — {o["type"]}, v{o["ver"]}{state}'
               f' · [открыть в ЛОЦМАН]({make_link(oid)}) · `{dirname}`')
    for dobj in [x for x in objects.values() if x['kind'] == 'doc' and x['parent'] == oid]:
        pad2 = '  ' * (depth + 1)
        out.append(f'{pad2}- _{dobj["type"]}_ **{dobj["product"]}** v{dobj["ver"]} ({dobj["state"]})'
                   f' · [открыть]({make_link(dobj["id"])})')
        for fl in docs.get(dobj['id'], []):
            dl = '✓' if fetched.get(fl['name'] + '|' + fl['doc_id']) else '·'
            out.append(f'{pad2}  - {dl} `{fl["name"]}` — {int(fl["size"] or 0) // 1024} КБ,'
                       f' CRC {fl["crc"]}, изм. {fl["modified"]}')

children = {}
for o in objects.values():
    if o['kind'] in ('obj', 'root'):
        children.setdefault(o['parent'], []).append(o['id'])

def walk(parent, depth, out, seen=None):
    seen = set() if seen is None else seen
    for cid in children.get(parent, []):
        if cid in seen:
            continue
        seen.add(cid)
        render_obj(cid, depth, out)
        walk(cid, depth + 1, out, seen)

walk(ROOT_ID, 0, lines)
for dobj in objects.values():
    if dobj['kind'] == 'doc' and dobj['parent'] not in objects:
        render_obj(dobj['parent'], 0, lines)

with open(os.path.join(OUT, 'index.md'), 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

# ---- .url в папках объектов (у объектов с файлами)
made = set()
for fl in files:
    owner = objects.get(fl['doc_id'], {}).get('parent', '')
    if not owner or owner == ROOT_ID:
        d = OUT
    else:
        d = os.path.join(OUT, rel_dir(owner))
    if not os.path.isdir(d):
        continue
    if owner in made:
        continue
    made.add(owner)
    o = objects.get(owner, {})
    nm = display(o) if o else owner
    with open(os.path.join(d, safe(nm) + ' - Loodsman.url'), 'w', encoding='cp1251') as f:
        f.write('[InternetShortcut]\r\nURL=' + make_link(owner) + '\r\n')

print(f'index.md, versions.json, .url x{len(made)} готово; '
      f'объектов {len(objects)}, файлов {len(files)}')
