# Loodsman-Skill

Скриптовое чтение ЛОЦМАН:PLM (АСКОН) 2014: обход дерева проекта,
выгрузка документов без взятия в работу, контроль версий по CRC,
индекс с гиперссылками `ask:Loodsman.URL`. Только чтение; канал —
официальный COM `DataProvider` (сервер приложений), клиент ЛОЦМАН
не требуется.

## Быстрый старт

1. Заполнить `local/loodsman.ini` и `local/config.json` (шаблоны рядом).
2. `python scripts/check_versions.py` — скан дерева и сравнение версий.
3. `cscript //Nologo scripts/loodsman_export.vbs fetch <tsvDir> <outDir> "" "pdf,docx"`
   — выгрузка с фильтром расширений.
4. `python scripts/build_index.py` — index.md + versions.json + .url-ссылки.

Подробности — `SKILL.md` и база `knowledge/`.

Лицензия: см. LICENSE.
