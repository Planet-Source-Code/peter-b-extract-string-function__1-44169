<div align="center">

## Extract String Function


</div>

### Description

This code will search a string for a given starting point(strFind) and a given end point(strSentinel) and return all the text in between. You can also add to the length of the start point if you don't want to include a number of unknown characters in the result
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Peter B](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/peter-b.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/peter-b-extract-string-function__1-44169/archive/master.zip)





### Source Code

```
Function ExtractString(ByVal strText As String, _
               strFind As String, _
               intAddToLen As Integer, _
               strSentinel As String, _
               TrapErrors As Boolean, _
               intLength As Integer) As String
  Dim SStart     As Integer
  Dim SEnd      As Integer
  SStart = InStr(1, strText, strFind) + Len(strFind) + intAddToLen
  If SStart <= Len(strFind) And TrapErrors = True Then
    MsgBox """" & strFind & """ not found!", vbCritical, "Error"
    Exit Function
  End If
  SEnd = InStr(SStart, strText, strSentinel)
  If SEnd <= Len(strFind) And TrapErrors = True Then
    MsgBox "Sentinel value """ & strSentinel & """ not found!", vbCritical, "Error"
    Exit Function
  End If
  If intLength > 0 Then
    ExtractString = Mid(strText, SStart, intLength)
  Else
    ExtractString = Mid(strText, SStart, (SEnd - SStart))
  End If
End Function
```

