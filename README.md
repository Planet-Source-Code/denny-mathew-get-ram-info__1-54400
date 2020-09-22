<div align="center">

## Get RAM Info


</div>

### Description

Gets RAM Info
 
### More Info
 
Gets RAM Info using WINAPI (GlobalMemoryStatus)

RAM Info


<span>             |<span>
---                |---
**Submitted On**   |2004-05-16 00:21:34
**By**             |[Denny Mathew](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/denny-mathew.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Get\_RAM\_In1757676152004\.zip](https://github.com/Planet-Source-Code/denny-mathew-get-ram-info__1-54400/archive/master.zip)

### API Declarations

```
Private Type MEMORYSTATUS
  dwLength As Long
  dwMemoryLoad As Long
 dwTotalPhys As Long
 dwAvailPhys As Long
 dwTotalPageFile As Long
 dwAvailPageFile As Long
 dwTotalVirtual As Long
 dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
```





