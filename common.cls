VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@func 读取文本文件内容,文件格式为utf-8
'@author ofix
'@date 2021/01/29
Public Static Function ReadTxtFile(path As String) As String
  If IsFileExists(path) = False Then
    ReadTxtFile = False
  End If
   With CreateObject("ADODB.Stream")
        .Type = 2    '读取
        .Mode = 3 '读写模式
        .Open
        .LoadFromFile path
        .Charset = "UTF-8" '编码格式
        .Position = 2
        ReadTxtFile = .ReadText
        .Close
    End With
End Function

'@func 判断文件是否存在
'@author ofix
'@date 2021/01/29
Public Static Function IsFileExists(ByVal path As String) As Boolean
 If Dir(path, 16) <> Empty Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function

'@func 高亮文本
'@author ofix
'@date 2021/01/29
Public Static Sub HilightWords(text As String)
   Dim textLen As Integer
   Dim i As Integer
   Dim j As Integer
   Dim word As String
   Dim posArr As New Collection '需要进行高亮的字符位置数组
   Dim translateArr As New Collection '转义字符位置
   textLen = Len(text)
   
   For i = 1 To textLen
     word = Mid(text, i, 1)
     Debug.Print "word=" & word
     If word = "@" Then
        posArr.Add i '添加当前位置元素
     ElseIf word = "\\" Then
        translateArr.Add i
        i = i + 1
     End If
   Next
   If posArr.Count Mod 2 <> 0 Then
    Debug.Print "数据长度不为偶数 " & posArr.Count
   End If
   '数组的数据
   For j = 0 To posArr.Count
    Debug.Print "pos = " & j
   Next
End Sub
