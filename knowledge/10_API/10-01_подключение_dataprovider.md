---
id: 10-01
title: Подключение к Лоцман из скрипта — DataProvider.LoodsmanConnection → API8 → SimpleAPI
tags: [api, подключение, dataprovider, com]
applies_to: ЛОЦМАН:PLM 2014, Windows x64 (нужен 32-битный cscript)
status: проверено
verified: 2026-10-30 — реальная выгрузка проекта: подключение, скан дерева, скачивание файлов
source: эксперимент + forum.ascon.ru (темы 32076, 34840)
---

# Подключение: DataProvider → API8 → SimpleAPI

Канал тот же, что у интеграторов АСКОН. Клиент ЛОЦМАН не нужен.

```vbs
Set cnn  = CreateObject("DataProvider.LoodsmanConnection")
cnn.Connected = True                     ' подключение к серверу приложений
Set api  = cnn.API8                      ' ILoodsmanAPI8
Set sapi = api.GetSimpleAPI()            ' ISimpleAPI
r = sapi.RunMethod("ConnectToDB", Array("ИмяБазы"))   ' 0 = успех
' чтение данных:
Set ds = CreateObject("DataProvider.DataSet")
ds.DATA = sapi.RunMethod("Метод", Array(параметры...))
ds.First : While Not ds.Eof : ds.FieldValue("_NAME") : ds.Next : Wend
```

## Ключевые факты

- `DataProvider.dll` (ProgID `DataProvider.LoodsmanAPI8`, `DataProvider.LoodsmanConnection`,
  `DataProvider.DataSet`) лежит в `Program Files (x86)\Common Files\ASCON Shared\Loodsman\`
  и ставится вместе с клиентом; это **32-битный inproc COM** — запускать только
  из 32-битного процесса: `C:\Windows\SysWOW64\cscript.exe` (или 32-битный Python).
  Из 64-битного Python напрямую — «Операция недоступна».
- Адрес сервера приложений не передаётся: `Connected = True` берёт его из
  настроек машины (реестр `HKCU\Software\ASCON\Loodsman`, `SP`). Ошибка
  «Невозможно активизировать методы сервера приложений» = сервер недоступен
  или имя в реестре устарело.
- Подключение занимает лицензию Лоцман. Освобождать: `Set sapi = Nothing`,
  `Set api = Nothing`, `cnn.Connected = False`, `Set cnn = Nothing`.
- Два интерфейса запуска методов: `sapi.RunMethod(имя, массив)` → VARIANT
  (датасет в сыром виде — присвоить `ds.DATA`), и `api.GetDataSet(имя, массив)`
  → готовый IDataSet. Оба проверены.
- `ConnectToDB` возвращает 0 при успехе; при неверном имени базы сервер
  показывает диалог «Укажите способ соединения с базой данных».

## Документация

SDK-справка `LoodsmanClientApi.chm` в поставку клиента не входит.
Полные типизированные сигнатуры всех серверных методов (440+) извлекаются
из библиотеки типов `AppServer\LoodsmanServerApplication.dll` — см.
`knowledge/10_API/loodsman_imainsystem_methods.txt` (безотносительно
конкретной установки; ProgID на клиентской машине не зарегистрирован,
но интерфейс `IMainSystem` — это то, что исполняет сервер приложений).
