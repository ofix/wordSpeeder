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

'@func ��ȡ�ı��ļ�����,�ļ���ʽΪutf-8
'@author ofix
'@date 2021/01/29
Public Static Function ReadTxtFile(path As String) As String
  If IsFileExists(path) = False Then
    ReadTxtFile = False
  End If
   With CreateObject("ADODB.Stream")
        .Type = 2    '��ȡ
        .Mode = 3 '��дģʽ
        .Open
        .LoadFromFile path
        .Charset = "UTF-8" '�����ʽ
        .Position = 2
        ReadTxtFile = .ReadText
        .Close
    End With
End Function

'@func �ж��ļ��Ƿ����
'@author ofix
'@date 2021/01/29
Public Static Function IsFileExists(ByVal path As String) As Boolean
 If Dir(path, 16) <> Empty Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function

'@func �����ı�
'@author ofix
'@date 2021/01/29
Public Static Sub HilightWords(text As String)
   Dim textLen As Integer
   Dim i As Integer
   Dim j As Integer
   Dim word As String
   Dim posArr As New Collection '��Ҫ���и������ַ�λ������
   Dim translateArr As New Collection 'ת���ַ�λ��
   textLen = Len(text)
   
   For i = 1 To textLen
     word = Mid(text, i, 1)
     Debug.Print "word=" & word
     If word = "@" Then
        posArr.Add i '���ӵ�ǰλ��Ԫ��
     ElseIf word = "\\" Then
        translateArr.Add i
        i = i + 1
     End If
   Next
   If posArr.Count Mod 2 <> 0 Then
    Debug.Print "���ݳ��Ȳ�Ϊż�� " & posArr.Count
   End If
   '���������
   For j = 0 To posArr.Count
    Debug.Print "pos = " & j
   Next
End Sub