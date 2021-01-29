VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xWarnBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_warnBoxLeft As Single 'xWarnBox���
Private m_warnBoxTop As Single 'xWarnBox�ұ�
Private m_logo As String 'logo·��
Private m_wLogo As Single 'logo����
Private m_hLogo As Single 'logo�߶�
Private m_strOrigin As String '�ı�������
Private m_strFormat As String '��ʽ���������
Private m_textTitle '��������
Private m_padding As Single '�Ի����ڱ߾�
Private m_verticalMargin As Single '�Ի������¼��
Private m_fontFamily As String '����
Private m_titleFontFamily As String '��������
Private m_fontSize As Single '�����С
Private m_clrBk As Long '����ɫ
Private m_clrText1 As Long 'ǰ��ɫ1
Private m_clrText2 As Long 'ǰ��ɫ2
Private m_shapeBackgroundRect As Shape 'Բ�Ǿ���
Private m_picLogo As Shape 'logo
Private m_textboxTitle As Shape '�����ı���
Private m_textboxContent As Shape '�����ı���
Private m_warnBox '���ͼ��



'Logo·��
Public Property Get Logo() As String
   Logo = m_logo
End Property

'Logo����
Public Property Get LogoWidth() As String
  LogoWidth = m_wLogo
End Property

'Logo�߶�
Public Property Get LogoHeight() As String
  LogoHeight = m_hLogo
End Property

'��������
Public Property Get TextTitle() As String
   TextTitle = m_textTitle
End Property

'��������
Public Property Get TitleFontFamily() As String
   TitleFontFamily = m_titleFontFamily
End Property

'���ֿ�����
Public Property Get text() As String
   text = m_strOrigin
End Property

'��ʽ������ı�������
Public Property Get TextFormat() As String
   TextFormat = m_strFormat
End Property

'�ı����������������¼��
Public Property Get VerticalMargin() As Single
   VerticalMargin = m_verticalMargin
End Property

'��������
Public Property Get FontFamily() As String
   FontFamily = m_fontFamily
End Property
'�����С
Public Property Get FontSize() As Single
   FontSize = m_fontSize
End Property




Public Property Let LogoWidth(w As String)
   m_wLogo = w
End Property

Property Let LogoHeight(h As String)
   m_hLogo = h
End Property

Property Let Logo(path As String)
  m_logo = path
End Property

'��������
Public Property Let TextTitle(text As String)
    m_textTitle = text
End Property

'��������
Public Property Let TitleFontFamily(family As String)
   m_titleFontFamily = family
End Property

'�����ı���
Public Property Let text(str As String)
   m_strOrigin = str
End Property

'��ʽ�������ı���
Public Property Let TextFormat(str As String)
   m_strFormat = str
End Property

'��ֱ���
Public Property Let VerticalMargin(margin As Single)
   m_verticalMargin = margin
End Property

'����
Public Property Let FontFamily(family As String)
   m_fontFamily = family
End Property

'�����С
Public Property Let FontSize(size As Single)
  m_fontSize = size
End Property



'@func ���Ʊ���Բ�Ǿ���
'@author ofix
'@date 2021/01/28
Private Sub DrawBackgroundBox()
  Dim docWidth As Single
  docWidth = ActiveDocument.PageSetup.PageWidth - ActiveDocument.PageSetup.LeftMargin _
                        - ActiveDocument.PageSetup.RightMargin
  '����Բ�Ǿ��α���ͼ
  Set m_shapeBackgroundRect = ActiveDocument.Shapes.AddShape(msoShapeRoundedRectangle, m_warnBoxLeft, m_warnBoxTop, docWidth, 50)
  With m_shapeBackgroundRect
    .Fill.ForeColor = m_clrText2 '��������ɫΪ rgb(226,218,241)
    .Line.Visible = msoFalse '�Ƴ��߿�
    .Adjustments.Item(1) = 0.04 '����Բ�Ǿ���
    .ZOrder msoSendToBack '����һ��
  End With
End Sub



'@func ����Logo
'@author ofix
'@date 2021/01/28
Private Sub DrawLogo()   '����LogoͼƬ
  Dim picLeft As Single
  Dim picTop As Single
  picLeft = m_warnBoxLeft - ActiveDocument.PageSetup.LeftMargin + m_padding
  picTop = m_warnBoxTop - ActiveDocument.PageSetup.TopMargin + m_padding
  Set m_picLogo = ActiveDocument.Shapes.AddPicture(m_logo, False, True, _
                  picLeft, picTop, m_wLogo, m_hLogo)
End Sub



'@func ���ӱ�����
'@author ofix
'@date 2021/01/28
Private Sub DrawTitle()
  Dim titleLeft As Single
  Dim titleTop As Single
  Dim titleWidth As Single
  Dim titleHeight As Single
  Dim titleBox As Shape
  titleLeft = m_warnBoxLeft + m_padding + m_wLogo + 4
  titleTop = m_warnBoxTop + m_padding
  titleWidth = Len(m_textTitle) * (m_fontSize * 1.5)
  titleHeight = m_hLogo
  
  Set m_textboxTitle = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, titleLeft, titleTop, titleWidth, titleHeight)
  With m_textboxTitle
    .Line.Visible = msoFalse '���ر߿�
    .TextFrame.TextRange.text = m_textTitle '������������
    .TextFrame.marginLeft = 0  '�����ڱ߾�
    .TextFrame.marginRight = 0
    .TextFrame.MarginTop = 0
    .TextFrame.MarginBottom = 0
    .TextFrame.AutoSize = msoAutoSizeShapeToFitText '����Ӧ����
    .TextFrame.WordWrap = msoTrue
    .TextFrame.TextRange.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '�̶��о�Ϊ16��
    .TextFrame.TextRange.ParagraphFormat.LineSpacing = m_fontSize '��λ:��
    .TextFrame.TextRange.Font.Name = m_titleFontFamily
    .TextFrame.TextRange.Font.Bold = True '�Ӵ�
    .TextFrame.TextRange.Font.NameFarEast = m_fontFamily '��������
    .TextFrame.TextRange.Font.size = m_fontSize '���������С
    .TextFrame.TextRange.Font.Fill.ForeColor.RGB = m_clrText1 '�ı�������ɫ
  End With
End Sub



'@func �����ı�������
'@author ofix
'@date 2021/01/28
Private Sub DrawText()
 '����������
  Dim contentLeft As Single
  Dim contentTop As Single
  Dim contentWidth As Single
  contentLeft = m_warnBoxLeft + m_padding
  contentTop = m_warnBoxTop + m_padding + m_hLogo + m_verticalMargin '�ڱ߾�+logo�߶�+��ֱ���
  contentWidth = m_shapeBackgroundRect.Width - m_padding * 2
  Set m_textboxContent = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, contentLeft, contentTop, contentWidth, 20)
  With m_textboxContent
    .Line.Visible = msoFalse '���ر߿�
    .TextFrame.AutoSize = True '����Ӧ���֣���Word 2010��True����������Ч���°汾��Word����Ҫ�ĳ� msoAutoSizeShapeToFitText
    .TextFrame.TextRange.text = m_strFormat '������������
    .TextFrame.marginLeft = 0  '�����ڱ߾�
    .TextFrame.marginRight = 0
    .TextFrame.MarginTop = 0
    .TextFrame.MarginBottom = 0
    .TextFrame.TextRange.Font.Name = m_fontFamily
    .TextFrame.TextRange.Font.NameFarEast = m_fontFamily '��������
    .TextFrame.TextRange.Font.size = m_fontSize '���������С
    .TextFrame.TextRange.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '�̶��о�
    .TextFrame.TextRange.ParagraphFormat.LineSpacing = m_fontSize + 4
    .TextFrame.TextRange.Font.Fill.ForeColor.RGB = m_clrText1 '�ı�������ɫ
    .TextFrame.WordWrap = msoTrue
  End With
  'Debug.Print ">>>m_textBox = " & m_textboxContent.Height
End Sub


'@func ����ĵ�
'@author ofix
'@date 2021/01/28
Private Sub ClearDocument()
  Selection.WholeStory
  Selection.Delete Unit:=wdCharacter, Count:=1
End Sub



'@func �������Ƶ�λ��
'@author ofix
'@date 2021/01/28
Private Sub MoveToNextLine()
  Selection.EndKey Unit:=wdLine, Extend:=wdExtend
  Selection.InsertAfter vbCrLf
  Selection.InsertAfter vbCrLf
  Selection.MoveDown Unit:=wdLine, Count:=1
End Sub


'@func �������ͼ�β�Ƕ�뵽����
'@author ofix
'@date 2021/01/28
Private Sub GroupShapes()
  Set m_warnBox = ActiveDocument.Shapes.Range(Array(m_shapeBackgroundRect.Name, m_picLogo.Name, m_textboxTitle.Name, m_textboxContent.Name)).Group
  m_warnBox.WrapFormat.Type = wdWrapInline  '�������ͼ�����ֻ��Ʒ�ʽ��Ƕ���͡�
End Sub


'@func ���ñ������С�����ͼ��
'@author ofix
'@date 2021/01/28
Private Sub ResizeBackgroundBox()
     m_shapeBackgroundRect.Height = m_padding * 2 + m_hLogo + m_verticalMargin + m_textboxContent.Height
     'Debug.Print "m_padding = " & m_padding & " m_hLogo = " & m_hLogo & " m_vertialMargin = " & m_verticalMargin & " m_textBox = " & m_textboxContent.Height
End Sub

'@func ����WarnBox
'@author ofix
'@date 2021/01/28
Public Sub Draw()
  'ClearDocument
  DrawBackgroundBox
  DrawLogo
  DrawTitle
  DrawText
  ResizeBackgroundBox
  GroupShapes
  MoveToNextLine
End Sub

'@func xWarnBox��ʼ��
'@author ofix
'@date 2021/01/28
Private Sub Class_Initialize()
  m_warnBoxLeft = ActiveDocument.PageSetup.LeftMargin
  m_warnBoxTop = Selection.Information(wdVerticalPositionRelativeToPage) '��ȡ��ǰ�������λ�ô�����һ��ĵط�
  m_logo = ActiveDocument.path & "\light.png"
  m_wLogo = 10 'logo����
  m_hLogo = 10 'logo�߶�
  Dim bmcFile As String
  bmcFile = ActiveDocument.path & "\bmc.txt"
  Dim cm As New common
  m_strOrigin = cm.ReadTxtFile(bmcFile)
  cm.HilightWords (m_strOrigin)
  m_strFormat = m_strOrigin
  m_padding = 10
  m_verticalMargin = 10 '�Ի������¼��
  m_textTitle = "ע��"
  m_fontFamily = "Microsoft YaHei" '΢���ź�����
  m_titleFontFamily = "΢���ź�" '΢���ź�����
  m_fontSize = 10 '�����С
  m_clrText1 = RGB(56, 34, 126)
  m_clrText2 = RGB(226, 218, 241)
  Set m_shapeBackgroundRect = Nothing
  Set m_textboxTitle = Nothing
  Set m_textboxContent = Nothing
  Set m_picLogo = Nothing
  Set m_warnBox = Nothing
End Sub