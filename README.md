<div align="center">

## CopyFileAny


</div>

### Description

This code allows you to copy any file, including your application to another destination at runtime.

Note: This code has only been tested on Windows 98 using Visual Basic 6.0
 
### More Info
 
Input file, output file

Note: This sub is only useful for copying inaccessable files. For regular files, the "FileCopy" sub should be called. An example of use would be to make your exe copy itself somewhere else while running.

Returns 1 on success, 0 on error.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Owens](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-owens.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-owens-copyfileany__1-3324/archive/master.zip)





### Source Code

```
Public Function CopyFileAny(currentFilename As String, newFilename As String)
Dim a%, buffer%, temp$, fRead&, fSize&, b%
On Error GoTo ErrHan:
a = FreeFile
buffer = 4048
 Open currentFilename For Binary Access Read As a
 b = FreeFile
 Open newFilename For Binary Access Write As b
 fSize = FileLen(currentFilename)
 While fRead < fSize
 DoEvents
 If buffer > (fSize - fRead) Then buffer = (fSize - fRead)
 temp = Space(buffer)
 Get a, , temp
 Put b, , temp
 fRead = fRead + buffer
 Wend
 Close b
 Close a
CopyFileAny=1
Exit Function
ErrHan:
CopyFileAny=0
End Function
```

