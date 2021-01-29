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

Private m_warnBoxLeft As Single 'xWarnBox左边
Private m_warnBoxTop As Single 'xWarnBox右边
Private m_logo As String 'logo路径
Private m_wLogo As Single 'logo宽度
Private m_hLogo As Single 'logo高度
Private m_strOrigin As String '文本框内容
Private m_strFormat As String '格式化后的内容
Private m_textTitle '标题文字
Private m_padding As Single '对话框内边距
Private m_verticalMargin As Single '对话框上下间距
Private m_fontFamily As String '字体
Private m_titleFontFamily As String '标题字体
Private m_fontSize As Single '字体大小
Private m_clrBk As Long '背景色
Private m_clrText1 As Long '前景色1
Private m_clrText2 As Long '前景色2
Private m_shapeBackgroundRect As Shape '圆角矩形
Private m_picLogo As Shape 'logo
Private m_textboxTitle As Shape '标题文本框
Private m_textboxContent As Shape '内容文本框
Private m_warnBox '组合图形



'Logo路径
Public Property Get Logo() As String
   Logo = m_logo
End Property

'Logo宽度
Public Property Get LogoWidth() As String
  LogoWidth = m_wLogo
End Property

'Logo高度
Public Property Get LogoHeight() As String
  LogoHeight = m_hLogo
End Property

'标题文字
Public Property Get TextTitle() As String
   TextTitle = m_textTitle
End Property

'标题字体
Public Property Get TitleFontFamily() As String
   TitleFontFamily = m_titleFontFamily
End Property

'文字框文字
Public Property Get text() As String
   text = m_strOrigin
End Property

'格式化后的文本框文字
Public Property Get TextFormat() As String
   TextFormat = m_strFormat
End Property

'文本框距离标题栏的上下间距
Public Property Get VerticalMargin() As Single
   VerticalMargin = m_verticalMargin
End Property

'字体名称
Public Property Get FontFamily() As String
   FontFamily = m_fontFamily
End Property
'字体大小
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

'标题文字
Public Property Let TextTitle(text As String)
    m_textTitle = text
End Property

'标题字体
Public Property Let TitleFontFamily(family As String)
   m_titleFontFamily = family
End Property

'内容文本框
Public Property Let text(str As String)
   m_strOrigin = str
End Property

'格式化内容文本框
Public Property Let TextFormat(str As String)
   m_strFormat = str
End Property

'垂直间距
Public Property Let VerticalMargin(margin As Single)
   m_verticalMargin = margin
End Property

'字体
Public Property Let FontFamily(family As String)
   m_fontFamily = family
End Property

'字体大小
Public Property Let FontSize(size As Single)
  m_fontSize = size
End Property



'@func 绘制背景圆角矩形
'@author ofix
'@date 2021/01/28
Private Sub DrawBackgroundBox()
  Dim docWidth As Single
  docWidth = ActiveDocument.PageSetup.PageWidth - ActiveDocument.PageSetup.LeftMargin _
                        - ActiveDocument.PageSetup.RightMargin
  '添加圆角矩形背景图
  Set m_shapeBackgroundRect = ActiveDocument.Shapes.AddShape(msoShapeRoundedRectangle, m_warnBoxLeft, m_warnBoxTop, docWidth, 50)
  With m_shapeBackgroundRect
    .Fill.ForeColor = m_clrText2 '调整背景色为 rgb(226,218,241)
    .Line.Visible = msoFalse '移除边框
    .Adjustments.Item(1) = 0.04 '设置圆角矩形
    .ZOrder msoSendToBack '上移一层
  End With
End Sub



'@func 绘制Logo
'@author ofix
'@date 2021/01/28
Private Sub DrawLogo()   '插入Logo图片
  Dim picLeft As Single
  Dim picTop As Single
  picLeft = m_warnBoxLeft - ActiveDocument.PageSetup.LeftMargin + m_padding
  picTop = m_warnBoxTop - ActiveDocument.PageSetup.TopMargin + m_padding
  Set m_picLogo = ActiveDocument.Shapes.AddPicture(m_logo, False, True, _
                  picLeft, picTop, m_wLogo, m_hLogo)
End Sub



'@func 添加标题栏
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
    .Line.Visible = msoFalse '隐藏边框
    .TextFrame.TextRange.text = m_textTitle '设置文字内容
    .TextFrame.marginLeft = 0  '调整内边距
    .TextFrame.marginRight = 0
    .TextFrame.MarginTop = 0
    .TextFrame.MarginBottom = 0
    .TextFrame.AutoSize = msoAutoSizeShapeToFitText '自适应文字
    .TextFrame.WordWrap = msoTrue
    .TextFrame.TextRange.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行距为16磅
    .TextFrame.TextRange.ParagraphFormat.LineSpacing = m_fontSize '单位:磅
    .TextFrame.TextRange.Font.Name = m_titleFontFamily
    .TextFrame.TextRange.Font.Bold = True '加粗
    .TextFrame.TextRange.Font.NameFarEast = m_fontFamily '设置字体
    .TextFrame.TextRange.Font.size = m_fontSize '设置字体大小
    .TextFrame.TextRange.Font.Fill.ForeColor.RGB = m_clrText1 '改变文字颜色
  End With
End Sub



'@func 添加文本框内容
'@author ofix
'@date 2021/01/28
Private Sub DrawText()
 '添加内容区
  Dim contentLeft As Single
  Dim contentTop As Single
  Dim contentWidth As Single
  contentLeft = m_warnBoxLeft + m_padding
  contentTop = m_warnBoxTop + m_padding + m_hLogo + m_verticalMargin '内边距+logo高度+垂直间距
  contentWidth = m_shapeBackgroundRect.Width - m_padding * 2
  Set m_textboxContent = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, contentLeft, contentTop, contentWidth, 20)
  With m_textboxContent
    .Line.Visible = msoFalse '隐藏边框
    .TextFrame.AutoSize = True '自适应文字，在Word 2010中True可以设置生效，新版本的Word可能要改成 msoAutoSizeShapeToFitText
    .TextFrame.TextRange.text = m_strFormat '设置文字内容
    .TextFrame.marginLeft = 0  '调整内边距
    .TextFrame.marginRight = 0
    .TextFrame.MarginTop = 0
    .TextFrame.MarginBottom = 0
    .TextFrame.TextRange.Font.Name = m_fontFamily
    .TextFrame.TextRange.Font.NameFarEast = m_fontFamily '设置字体
    .TextFrame.TextRange.Font.size = m_fontSize '设置字体大小
    .TextFrame.TextRange.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行距
    .TextFrame.TextRange.ParagraphFormat.LineSpacing = m_fontSize + 4
    .TextFrame.TextRange.Font.Fill.ForeColor.RGB = m_clrText1 '改变文字颜色
    .TextFrame.WordWrap = msoTrue
  End With
  'Debug.Print ">>>m_textBox = " & m_textboxContent.Height
End Sub


'@func 清空文档
'@author ofix
'@date 2021/01/28
Private Sub ClearDocument()
  Selection.WholeStory
  Selection.Delete Unit:=wdCharacter, Count:=1
End Sub



'@func 增加下移的位置
'@author ofix
'@date 2021/01/28
Private Sub MoveToNextLine()
  Selection.EndKey Unit:=wdLine, Extend:=wdExtend
  Selection.InsertAfter vbCrLf
  Selection.InsertAfter vbCrLf
  Selection.MoveDown Unit:=wdLine, Count:=1
End Sub


'@func 组合所有图形并嵌入到文字
'@author ofix
'@date 2021/01/28
Private Sub GroupShapes()
  Set m_warnBox = ActiveDocument.Shapes.Range(Array(m_shapeBackgroundRect.Name, m_picLogo.Name, m_textboxTitle.Name, m_textboxContent.Name)).Group
  m_warnBox.WrapFormat.Type = wdWrapInline  '设置组合图形文字环绕方式，嵌入型。
End Sub


'@func 重置背景框大小并组合图形
'@author ofix
'@date 2021/01/28
Private Sub ResizeBackgroundBox()
     m_shapeBackgroundRect.Height = m_padding * 2 + m_hLogo + m_verticalMargin + m_textboxContent.Height
     'Debug.Print "m_padding = " & m_padding & " m_hLogo = " & m_hLogo & " m_vertialMargin = " & m_verticalMargin & " m_textBox = " & m_textboxContent.Height
End Sub

'@func 绘制WarnBox
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

'@func xWarnBox初始化
'@author ofix
'@date 2021/01/28
Private Sub Class_Initialize()
  m_warnBoxLeft = ActiveDocument.PageSetup.LeftMargin
  m_warnBoxTop = Selection.Information(wdVerticalPositionRelativeToPage) '获取当前光标所在位置处于上一层的地方
  m_logo = ActiveDocument.path & "\light.png"
  m_wLogo = 10 'logo宽度
  m_hLogo = 10 'logo高度
  Dim bmcFile As String
  bmcFile = ActiveDocument.path & "\bmc.txt"
  Dim cm As New common
  m_strOrigin = cm.ReadTxtFile(bmcFile)
  cm.HilightWords (m_strOrigin)
  m_strFormat = m_strOrigin
  m_padding = 10
  m_verticalMargin = 10 '对话框上下间距
  m_textTitle = "注意"
  m_fontFamily = "Microsoft YaHei" '微软雅黑字体
  m_titleFontFamily = "微软雅黑" '微软雅黑字体
  m_fontSize = 10 '字体大小
  m_clrText1 = RGB(56, 34, 126)
  m_clrText2 = RGB(226, 218, 241)
  Set m_shapeBackgroundRect = Nothing
  Set m_textboxTitle = Nothing
  Set m_textboxContent = Nothing
  Set m_picLogo = Nothing
  Set m_warnBox = Nothing
End Sub
