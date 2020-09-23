<div align="center">

## CurComputerName


</div>

### Description

Get the current computername from Windows
 
### More Info
 
sTmp1$ = CurComputerName()

( The current computername from Windows )

sTmp1$ = "MORTAL OBSESSiON"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MAGiC MANiAC^mTo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/magic-maniac-mto.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/magic-maniac-mto-curcomputername__1-2338/archive/master.zip)





### Source Code

```
' Get the current computername from Windows
' Coded By MAGiC MANiAC^mTo
' More Examples At: http://home.kabelfoon.nl/~mto/
'
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _ (ByVal lpBuffer$, nSize As Long) As Long
Function CurComputerName$()
 Dim sTmp1$
 sTmp1 = Space$(512)
 GetComputerName sTmp1, Len(sTmp1)
 CurComputerName = Trim$(sTmp1)
End Function
```

