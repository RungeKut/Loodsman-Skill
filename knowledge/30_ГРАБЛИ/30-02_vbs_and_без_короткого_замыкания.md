---
id: 30-02
title: В VBScript And/Or вычисляют оба операнда — проверки вида If X And F(X) падают
tags: [грабли, vbs, логика]
applies_to: VBScript (WSH)
status: проверено
verified: 2026-10-30 — If fso.FileExists(dst) And fso.GetFile(dst).Size=… падало «Файл не найден» на отсутствии файла
source: эксперимент
---

# And в VBScript не короткозамкнутый

## Что происходит

`If fso.FileExists(p) And fso.GetFile(p).Size = n` выполняет `GetFile`
даже когда файла нет → «Файл не найден». То же с `InStr(s,"|") > 0 And
Split(s,"|")(1)` — Split по строке без разделителя даёт «Индекс выходит
за пределы диапазона».

## Как правильно

Вложенные If или предварительное вычисление в переменную:

```vbs
bHave = False
If fso.FileExists(dst) Then
  If size <> "" Then
    If fso.GetFile(dst).Size = CLng(size) Then bHave = True
  End If
End If
```

## Как неправильно

Полагаться на «умное» сокращение как в C/Python — в VBScript его нет.

## Чем подтверждено

Две аварии в `loodsman_export.vbs` за один вечер, обе на этом.
