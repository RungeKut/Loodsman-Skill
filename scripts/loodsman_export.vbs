' ============================================================
' loodsman_export.vbs - scan/fetch of a Loodsman PLM project.
' Read-only API usage: GetTree, GetInfoAboutVersion(mode 7),
' GetLinkedFast2, GetFileById. No writes to the database.
'
' Usage:
'   cscript //Nologo loodsman_export.vbs scan  <rootId> <tsvDir>
'   cscript //Nologo loodsman_export.vbs fetch <tsvDir> <outDir> [excludeExts] [includeExts]
'
' Config (ANSI/cp1251 ini, not under git):
'   <skillRoot>\local\loodsman.ini
'     db=<database display name>
'     linkComposition=<composition link type name>
'     linkDocuments=<documents link type name>
'   32-bit cscript required (SysWOW64) - DataProvider.dll is x86 COM.
' TSV files: cp1251, TAB-separated.
' ============================================================
Option Explicit

Dim gIni, gFso
Set gFso = CreateObject("Scripting.FileSystemObject")

Function ScriptDir()
  ScriptDir = gFso.GetParentFolderName(WScript.ScriptFullName)
End Function

Function SkillRoot()
  SkillRoot = gFso.GetParentFolderName(ScriptDir())
End Function

' ini: single "key=value" per line, ANSI encoding
Function IniValue(key)
  Dim p, f, line, k
  IniValue = ""
  p = SkillRoot() & "\local\loodsman.ini"
  If Not gFso.FileExists(p) Then Exit Function
  Set f = gFso.OpenTextFile(p, 1, False, False)
  Do While Not f.AtEndOfStream
    line = f.ReadLine
    k = Split(line & "=", "=")(0)
    If Trim(k) = key Then
      IniValue = Trim(Mid(line, InStr(line, "=") + 1))
      Exit Do
    End If
  Loop
  f.Close
End Function

Dim DB_NAME, LINK_COMP, LINK_DOCS
DB_NAME = IniValue("db")
LINK_COMP = IniValue("linkComposition")
LINK_DOCS = IniValue("linkDocuments")
If DB_NAME = "" Or LINK_COMP = "" Or LINK_DOCS = "" Then
  WScript.Echo "ERROR: fill " & SkillRoot() & "\local\loodsman.ini" & _
    " with db=, linkComposition=, linkDocuments= (ANSI encoding)"
  WScript.Quit 1
End If

Dim gCnn, gApi, gSapi
Set gFso = CreateObject("Scripting.FileSystemObject")

Function Connect()
  Set gCnn = CreateObject("DataProvider.LoodsmanConnection")
  gCnn.Connected = True
  If Not gCnn.Connected Then
    WScript.Echo "ERROR: no app server connection"
    WScript.Quit 1
  End If
  Set gApi = gCnn.API8
  Set gSapi = gApi.GetSimpleAPI()
  Dim r
  r = gSapi.RunMethod("ConnectToDB", Array(DB_NAME))
  If r <> 0 Then
    WScript.Echo "ERROR: ConnectToDB returned " & r
    WScript.Quit 1
  End If
End Function

Sub Disconnect()
  Set gSapi = Nothing
  Set gApi = Nothing
  gCnn.Connected = False
  Set gCnn = Nothing
End Sub

Function NewDS()
  Set NewDS = CreateObject("DataProvider.DataSet")
End Function

Function RunData(method, params)
  Dim ds
  Set ds = NewDS()
  On Error Resume Next
  ds.DATA = gSapi.RunMethod(method, params)
  If Err.Number <> 0 Then
    WScript.Echo "WARN: " & method & " err=" & Err.Number & " " & Err.Description
    Err.Clear
    On Error GoTo 0
    Set RunData = Nothing
    Exit Function
  End If
  On Error GoTo 0
  Set RunData = ds
End Function

Function FV(ds, name)
  On Error Resume Next
  FV = ds.FieldValue(name)
  If Err.Number <> 0 Then
    Err.Clear
    FV = ""
  End If
  On Error GoTo 0
  If IsNull(FV) Then FV = ""
  FV = CStr(FV)
End Function

Function SafeName(s)
  Dim bad, i, c, r
  bad = "\/:*?""<>|" & vbTab
  r = ""
  For i = 1 To Len(s)
    c = Mid(s, i, 1)
    If InStr(bad, c) > 0 Then
      r = r & "_"
    Else
      r = r & c
    End If
  Next
  If Len(r) > 120 Then r = Left(r, 120)
  SafeName = r
End Function

Sub MakeDirMulti(path)
  Dim parts, cur, i
  If gFso.FolderExists(path) Then Exit Sub
  parts = Split(path, "\")
  cur = parts(0)
  For i = 1 To UBound(parts)
    cur = cur & "\" & parts(i)
    If Not gFso.FolderExists(cur) Then gFso.CreateFolder(cur)
  Next
End Sub

' ---------- SCAN ----------
' gObjects: id -> id|kind|parent|type|product|version|state|link
' kind: root|obj|doc ; markers are string keys "D<id>" -> "done"
Dim gObjects, gFiles
Set gObjects = CreateObject("Scripting.Dictionary")
Set gFiles = CreateObject("Scripting.Dictionary") ' fileId -> TSV row

Sub ScanObject(idVersion, parentVersion, depth)
  Dim ds, row, childId, docId
  If gObjects.Exists("D" & idVersion) Then Exit Sub
  gObjects.Add "D" & idVersion, "done"

  Set ds = RunData("GetTree", Array("", "", "", CLng(idVersion), LINK_COMP, False))
  If Not ds Is Nothing Then
    ds.First
    Do While Not ds.Eof
      childId = CLng(FV(ds, "_ID_VERSION"))
      row = childId & "|obj|" & parentVersion & "|" & FV(ds, "_TYPE") & "|" & _
            FV(ds, "_PRODUCT") & "|" & FV(ds, "_VERSION") & "|" & _
            FV(ds, "_STATE") & "|" & FV(ds, "_LINKTYPE")
      If Not gObjects.Exists(childId) Then gObjects.Add childId, row
      ds.Next
    Loop
  End If

  Set ds = RunData("GetTree", Array("", "", "", CLng(idVersion), LINK_DOCS, False))
  If Not ds Is Nothing Then
    ds.First
    Do While Not ds.Eof
      docId = CLng(FV(ds, "_ID_VERSION"))
      If Not gObjects.Exists(docId) Then
        row = docId & "|doc|" & idVersion & "|" & FV(ds, "_TYPE") & "|" & _
              FV(ds, "_PRODUCT") & "|" & FV(ds, "_VERSION") & "|" & _
              FV(ds, "_STATE") & "|" & FV(ds, "_LINKTYPE")
        gObjects.Add docId, row
      End If
      ds.Next
    Loop
  End If

  Dim keys, k
  keys = gObjects.Keys
  For Each k In keys
    row = gObjects(k)
    If VarType(k) <> vbString And InStr(row, "|") > 0 Then
      If Split(row, "|")(1) = "obj" And Not gObjects.Exists("D" & k) Then
        ScanObject k, k, depth + 1
      End If
    End If
  Next
End Sub

Sub ScanFiles()
  Dim keys, k, row, parts, ds, kindWord
  keys = gObjects.Keys
  For Each k In keys
    row = gObjects(k)
    kindWord = ""
    If VarType(k) <> vbString And InStr(row, "|") > 0 Then kindWord = Split(row, "|")(1)
    If kindWord = "doc" Then
      parts = Split(row, "|")
      Set ds = RunData("GetInfoAboutVersion", Array(parts(3), parts(4), parts(5), 0, CLng(7)))
      If Not ds Is Nothing Then
        ds.First
        Do While Not ds.Eof
          Dim fileId
          fileId = CLng(FV(ds, "_ID_FILE"))
          If Not gFiles.Exists(fileId) Then
            gFiles.Add fileId, k & vbTab & parts(3) & vbTab & parts(4) & vbTab & _
              parts(5) & vbTab & fileId & vbTab & FV(ds, "_NAME") & vbTab & _
              FV(ds, "_SIZE") & vbTab & FV(ds, "_CRC") & vbTab & FV(ds, "_MODIFIED")
          End If
          ds.Next
        Loop
      End If
    End If
  Next
End Sub

Function WriteTSV(path, header, dict)
  Dim f, keys, k
  Set f = gFso.CreateTextFile(path, True, False) ' ANSI
  f.WriteLine header
  keys = dict.Keys
  For Each k In keys
    f.WriteLine dict(k)
  Next
  f.Close
End Function

' ---------- FETCH ----------
Sub DoFetch(tsvDir, outDir, excl, incl)
  Dim f, line, parts, docId, fName, size, src, dst, docType, docProd, ext
  Dim nDone, nSkip, nErr, nSkipFilt, bHave, bSkip
  nDone = 0: nSkip = 0: nErr = 0: nSkipFilt = 0

  Set f = gFso.OpenTextFile(tsvDir & "\files.tsv", 1, False, False)
  f.SkipLine
  Dim outList
  Set outList = CreateObject("Scripting.Dictionary")

  Do While Not f.AtEndOfStream
    line = f.ReadLine
    parts = Split(line, vbTab)
    If UBound(parts) >= 8 Then
      docId = parts(0): docType = parts(1): docProd = parts(2)
      fName = parts(5): size = parts(6)
      If fName <> "" And Not outList.Exists(fName & "|" & docId) Then
        outList.Add fName & "|" & docId, line
        ext = LCase(gFso.GetExtensionName(fName))
        bSkip = False
        If incl <> "" Then
          If InStr(1, "," & incl & ",", "," & ext & ",", vbTextCompare) = 0 Then bSkip = True
        ElseIf excl <> "" And InStr(1, "," & excl & ",", "," & ext & ",", vbTextCompare) > 0 Then
          bSkip = True
        End If
        If bSkip Then
          nSkipFilt = nSkipFilt + 1
        Else
          Dim dstDir
          dstDir = outDir & "\" & SafeName(docProd) & " [" & SafeName(docType) & "]"
          MakeDirMulti dstDir
          dst = dstDir & "\" & SafeName(fName)

          bHave = False
          If gFso.FileExists(dst) Then
            If size <> "" Then
              If gFso.GetFile(dst).Size = CLng(size) Then bHave = True
            Else
              bHave = True
            End If
          End If
          If bHave Then
            nSkip = nSkip + 1
          Else
            On Error Resume Next
            src = gSapi.RunMethod("GetFileById", Array(CLng(docId), fName, ""))
            If Err.Number <> 0 Then
              WScript.Echo "ERR GetFileById [" & docProd & "] " & fName & ": " & Err.Description
              nErr = nErr + 1
              Err.Clear
            ElseIf src = "" Or Not gFso.FileExists(src) Then
              WScript.Echo "ERR no file: " & src
              nErr = nErr + 1
            Else
              gFso.CopyFile src, dst, True
              If size = "" Or gFso.GetFile(dst).Size = CLng(size) Then
                nDone = nDone + 1
              Else
                WScript.Echo "ERR size mismatch: " & fName
                nErr = nErr + 1
              End If
            End If
            On Error GoTo 0
          End If
        End If
      End If
    End If
  Loop
  f.Close

  Dim fo, keys, k
  Set fo = gFso.CreateTextFile(tsvDir & "\files_done.tsv", True, False)
  fo.WriteLine "docId" & vbTab & "docType" & vbTab & "docProduct" & vbTab & _
    "docVer" & vbTab & "fileId" & vbTab & "fileName" & vbTab & "size" & vbTab & _
    "crc" & vbTab & "modified" & vbTab & "status"
  keys = outList.Keys
  For Each k In keys
    fo.WriteLine outList(k) & vbTab & "fetched"
  Next
  fo.Close

  WScript.Echo "Downloaded: " & nDone & ", already: " & nSkip & _
    ", filtered: " & nSkipFilt & ", errors: " & nErr
End Sub

' ---------- MAIN ----------
Dim mode
If WScript.Arguments.Count < 2 Then
  WScript.Echo "Usage: loodsman_export.vbs scan <rootId> <tsvDir> | fetch <tsvDir> <outDir> [exclude] [include]"
  WScript.Quit 1
End If
mode = WScript.Arguments(0)

If mode = "scan" Then
  Dim rootId, tsvDir
  rootId = CLng(WScript.Arguments(1))
  tsvDir = WScript.Arguments(2)
  Connect
  gObjects.Add rootId, rootId & "|root|" & rootId & "|ROOT||||"
  ScanObject rootId, rootId, 0
  ScanFiles
  Disconnect

  Dim objsNoMark, keys, k
  Set objsNoMark = CreateObject("Scripting.Dictionary")
  keys = gObjects.Keys
  For Each k In keys
    If VarType(k) <> vbString Then objsNoMark.Add k, gObjects(k)
  Next

  WriteTSV tsvDir & "\objects.tsv", "id|kind|parent|type|product|version|state|linktype", objsNoMark
  WriteTSV tsvDir & "\files.tsv", "docId" & vbTab & "docType" & vbTab & "docProduct" & vbTab & _
    "docVer" & vbTab & "fileId" & vbTab & "fileName" & vbTab & "size" & vbTab & "crc" & vbTab & "modified", gFiles

  WScript.Echo "Objects: " & objsNoMark.Count & ", files: " & gFiles.Count
  Dim totalSize
  totalSize = 0
  keys = gFiles.Keys
  For Each k In keys
    totalSize = totalSize + CLng(Split(gFiles(k), vbTab)(6))
  Next
  WScript.Echo "Total size, MB: " & FormatNumber(totalSize / 1048576, 1)

ElseIf mode = "fetch" Then
  Dim ftsv, outd, excl, incl
  ftsv = WScript.Arguments(1)
  outd = WScript.Arguments(2)
  excl = ""
  incl = ""
  If WScript.Arguments.Count >= 4 Then excl = LCase(WScript.Arguments(3))
  If WScript.Arguments.Count >= 5 Then incl = LCase(WScript.Arguments(4))
  Connect
  DoFetch ftsv, outd, excl, incl
  Disconnect
  WScript.Echo "DONE"
Else
  WScript.Echo "Unknown mode: " & mode
  WScript.Quit 1
End If
