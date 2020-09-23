<div align="center">

## Auto Start with Windows \(fixed\)


</div>

### Description

Make your program start when windows starts. Short and easy.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dound](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dound.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dound-auto-start-with-windows-fixed__1-49509/archive/master.zip)

### API Declarations

```
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
```


### Source Code

```
Const REG_SZ = 1
Const HKEY_CURRENT_USER = &H80000001
Const REGKEY = "Software\Microsoft\Windows\CurrentVersion\Run"
Const KEY_WRITE = &H20006
Dim path As Long
'Tell windows to make Autopaper autostart with windows
If RegOpenKeyEx(HKEY_CURRENT_USER, REGKEY, 0, KEY_WRITE, path) Then Exit Sub
RegSetValueEx path, App.Title, 0, REG_SZ, ByVal App.path & "\startprog.exe", Len(App.path & "\programsfilename.exe")
    'DELETE AUTOSTART:
    'If RegOpenKeyEx(HKEY_CURRENT_USER, REGKEY, 0, KEY_WRITE, Path) Then Exit Sub
    'RegDeleteValue Path, App.Title
```

