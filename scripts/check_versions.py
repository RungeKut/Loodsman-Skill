# -*- coding: utf-8 -*-
"""
check_versions.py — проверка версий проекта Лоцман и догрузка изменений.

Алгоритм:
  1. loodsman_export.vbs scan — свежий files.tsv (метаданные + CRC, без скачивания).
  2. Сравнение с versions.json прошлого прохода: новые / изменившиеся / исчезнувшие.
  3. С ключом --fetch: перекачка только изменившихся.

Запуск:
  python check_versions.py [tsvDir] [projectDir] [rootId] [include] [--fetch]
Без аргументов значения берутся из local/config.json.

Каждый запуск сканирования берёт лицензию Лоцман на время скана.
"""
import json
import os
import subprocess
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(SCRIPT_DIR)
CONFIG = os.path.join(SKILL_ROOT, 'local', 'config.json')

cfg = {}
if os.path.exists(CONFIG):
    with open(CONFIG, encoding='utf-8') as f:
        cfg = json.load(f)


def arg(n, key, default=''):
    if len(sys.argv) > n and not sys.argv[n].startswith('--'):
        return sys.argv[n]
    return cfg.get(key, default)


tsv = arg(1, 'tsv_dir')
proj = arg(2, 'project_dir')
root = str(arg(3, 'root_id'))
include = arg(4, 'include')
do_fetch = '--fetch' in sys.argv

if not (tsv and proj and root):
    sys.exit('Заполните local/config.json (tsv_dir, project_dir, root_id, include) или '
             'передайте: tsvDir projectDir rootId include [--fetch]')

# 32-битный cscript обязателен: DataProvider.dll - x86 COM
cs32 = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'SysWOW64', 'cscript.exe')
cs = cs32 if os.path.exists(cs32) else os.path.join(
    os.environ.get('WINDIR', r'C:\Windows'), 'System32', 'cscript.exe')
export = os.path.join(SCRIPT_DIR, 'loodsman_export.vbs')


def run_export(args):
    r = subprocess.run([cs, '//Nologo', export] + args, capture_output=True)
    out = r.stdout.decode('cp866', errors='replace')
    if r.returncode != 0:
        print(out)
        sys.exit(f'loodsman_export.vbs failed: {args}')
    return out


def read_files_tsv():
    rows = []
    with open(os.path.join(tsv, 'files.tsv'), encoding='cp1251') as f:
        next(f)
        for line in f:
            c = line.rstrip('\r\n').split('\t')
            if len(c) >= 9:
                rows.append(dict(doc_id=c[0], doc_type=c[1], doc_prod=c[2], doc_ver=c[3],
                                 file_id=c[4], name=c[5], size=c[6], crc=c[7], modified=c[8]))
    return rows


def safe(s):
    for ch in '\\/:*?"<>|\t':
        s = s.replace(ch, '_')
    return s[:120]


print('1) скан...')
print(run_export(['scan', root, tsv]).strip())
new = read_files_tsv()

old_path = os.path.join(proj, 'versions.json')
old_files = {}
if os.path.exists(old_path):
    with open(old_path, encoding='utf-8') as f:
        old = json.load(f)
    for fl in old.get('files', []):
        old_files[fl['file_id']] = fl

added, changed, gone = [], [], []
for fl in new:
    o = old_files.get(fl['file_id'])
    if o is None:
        if os.path.exists(old_path):
            added.append(fl)
    elif o['crc'] != fl['crc'] or o['size'] != fl['size']:
        changed.append(fl)
if os.path.exists(old_path):
    new_ids = {fl['file_id'] for fl in new}
    gone = [o for fid, o in old_files.items() if fid not in new_ids]

print(f"2) сравнение: файлов сейчас {len(new)}, новых {len(added)}, "
      f"изменившихся {len(changed)}, исчезнувших {len(gone)}")
for fl in changed:
    print(f"   изменён: {fl['doc_prod']} / {fl['name']} "
          f"(crc {old_files[fl['file_id']]['crc']} -> {fl['crc']})")
for fl in added:
    print(f"   новый:   {fl['doc_prod']} / {fl['name']}")
for fl in gone:
    print(f"   исчез:   {fl['doc_prod']} / {fl['name']}")

if do_fetch and (added or changed):
    # fetch пропускает файлы с совпадающим размером, поэтому изменившиеся
    # (тот же размер, другой crc) удаляем локально для принудительной перекачки
    for fl in changed:
        local = os.path.join(proj, safe(fl['doc_prod']) + ' [' + safe(fl['doc_type']) + ']',
                             safe(fl['name']))
        if os.path.exists(local):
            os.remove(local)
    print('3) догрузка изменений...')
    out = run_export(['fetch', tsv, proj, '', include])
    print(out.strip())
else:
    print('(без --fetch догрузки нет)')

print('Готово. Для обновления index.md запустите build_index.py')
