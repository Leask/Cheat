Attribute VB_Name = "modConv"
Option Explicit

Public Function BytesToBstr(strBody, CodeBase)
Dim ObjStream
Set ObjStream = CreateObject("Adodb.Stream")
With ObjStream
.Type = 1
.Mode = 3
.Open
.Write strBody
.Position = 0
.Type = 2
.Charset = CodeBase
BytesToBstr = .ReadText
.Close
End With
Set ObjStream = Nothing
End Function
