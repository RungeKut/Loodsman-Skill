---
name: loodsman-skill
description: "Скриптовое чтение ЛОЦМАН:PLM (АСКОН) 2014: найти проект по ссылке ask:Loodsman.URL, обойти дерево, выгрузить документы без взятия в работу, проверить версии по CRC и догрузить изменения, собрать индекс со ссылками на объекты. Применяй всегда при словах: Лоцман, ЛОЦМАН, PLM, АСКОН, выгрузить из Лоцмана, ссылки на объекты."
---

# Loodsman-Skill

Корень набора — папка этого файла. База `knowledge/` (карта — `INDEX.md`).
Общие правила — соседний скилл skill-skill. Только чтение: никакого
`CheckOut`, никаких записей в базу.

**Первым делом:** `git -C <корень> pull --rebase`; затем проверить
настройки `local/loodsman.ini` (cp1251: db, имена связей) и
`local/config.json` (пути, root_id, include) — формат в knowledge 40-01.

## Быстрый старт

```bash
python scripts/check_versions.py                    # скан дерева + сравнение версий
cscript //Nologo scripts/loodsman_export.vbs fetch <tsvDir> <outDir> "" "pdf,docx,spw"
python scripts/build_index.py                       # index.md + versions.json + *.url
python scripts/check_versions.py --fetch            # догрузить изменившееся (по CRC)
```

## Железные правила

1. Только чтение: белый список методов — knowledge 10-02. Незнакомые
   методы не вызывать; сигнатуры — knowledge/10_API/loodsman_imainsystem_methods.txt.
2. Клиент ЛОЦМАН из скриптов не запускать и не трогать: COM-классы
   `Loodsman.*` роняют его (30-05). Внешний канал — DataProvider (10-01),
   32-битный cscript (30-01, 40-01).
3. `.vbs` — только ASCII-литералы; кириллица из ini в cp1251 (30-01).
4. В набор — ничего о рабочих проектах: базы, серверы, изделия, пути —
   только в `local/` (skill-skill 40-02); «файл не сохранен в базе» —
   не ошибка, см. 30-03.
5. Длинные scan/fetch держат лицензию Лоцман — предупреждать человека
   и не гонять в пиковые часы.

## Ссылка на объект

`ask:Loodsman.URL?Action=Navigate,params=base64(cp1251("0|<База>|<узел>|<idВерсии>"))`
— в адресной строке проводника открывает объект в клиенте; build_index.py
раскладывает такие ссылки по index.md и .url-файлам (20-01).
