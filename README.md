<div align="center">

## Squeeze


</div>

### Description

Removes extra spaces from given string. eg. Squeeze("^^too^^many^^^spaces^^")returns "too^many^spaces". NB.Read a carat ^ as a single space above. Planet Source code mangles multiple spaces in submitted text.
 
### More Info
 
A string, strText.

Doesn't strip tabs, CRLFs or any characters from text EXCEPT char 32.

A string = strText without spaces and the beginning or end or more than 1 consecutive space within.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Bambrick](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-bambrick.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-bambrick-squeeze__1-30842/archive/master.zip)





### Source Code

```
Public Function Squeeze (ByVal strText As String) As String
 Dim intPos As Integer
 intPos = InStr(1, strText, Chr(32) & Chr(32))
 If intPos = 0 Then
 Squeeze = Trim(strText)
 Else
 strText = _
 Left(strText, intPos - 1) & _
 Mid(strText, intPos + 1)
 Squeeze = Squeeze(strText)
 End If
End Function
```

