# CHANGELOG

## 2026-10-30 (2)

- Режим find: поиск объекта по обозначению (FindObjects, %маска%) с готовыми
  ссылками ask:Loodsman.URL; GetVersionList и атрибуты GetTree (boWithAttrs)
  проверены и внесены в 10-02.
- Выгрузка: папки объектов (Обозначение Наименование Примечание) с
  вложенностью состава вместо папок по типу документа; GetTree boWithAttrs=True.

## 2026-10-30

- Набор создан: подключение DataProvider → API8 → SimpleAPI (10-01),
  белый список методов чтения (10-02), цикл выгрузки проекта
  scan/fetch/build_index/check_versions (20-01).
- Грабли: cp1251 в .vbs (30-01), не-короткозамкнутый And (30-02),
  файлы без тела в базе (30-03), пустые таблицы при прямом SQL (30-04),
  падение клиента от клиентских COM-классов (30-05).
- Скрипты: scripts/loodsman_export.vbs (ASCII), scripts/build_index.py,
  scripts/check_versions.py. Настройки установки — в local/ (не под git).
- Полный дамп сигнатур серверного API (IMainSystem) —
  knowledge/10_API/loodsman_imainsystem_methods.txt.
