VERSION 5.00
Begin VB.UserControl LCDDisplay 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1065
   ScaleHeight     =   69
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   71
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Image PicDigit 
      Height          =   735
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "LCDDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum DigitSize
    Small = 0
    Large = 1
End Enum

Public Enum mhc_Appearance
  [3D] = 1
  Thin = 2
End Enum

Public Enum mhc_BorderStyle
  None = 0
  Etched = 1
  Raised = 2
  Sunken = 3
  Line = 4
End Enum

Enum LeadChar
    Zero = 0
    Blank = 1
End Enum

' private value-holders
Private m_Font                As StdFont
Private m_Appearance          As mhc_Appearance
Private m_BorderStyle         As mhc_BorderStyle
Private m_BorderColor         As OLE_COLOR

Dim m_DigitSize As DigitSize      'digit's size (0=small,1=large)
Dim m_DigitCount As Integer     'number of digits to display
Dim m_Value As Double           'Value to display
Dim m_FillChar As LeadChar       ' Leading Zeros or Blank

Public Event Error()

Private Function InitDisplay()
    Dim intLoop As Integer
    
    UserControl.Cls
    
    Rem Unload All but 1 of the image boxes
    For intLoop = PicDigit.UBound To 1 Step -1
        Unload PicDigit(intLoop)
    Next
    
    Rem Setup the 1st box
    PicDigit(0).Left = 4
    PicDigit(0).Top = 4
    If m_DigitSize = 0 Then
        PicDigit(0).Width = smImageWidth
        PicDigit(0).Height = smImageHeight
    Else
        PicDigit(0).Width = lgImageWidth
        PicDigit(0).Height = lgImageHeight
    End If
    
    'PicDigit(0).BorderStyle = 1
    
    UserControl.Cls
    
    Rem Now Add Enough Image Controls to Handle Number of Digits
    For intLoop = 1 To m_DigitCount - 1
        Load PicDigit(intLoop)
        PicDigit(intLoop).Visible = True
        PicDigit(intLoop).Top = PicDigit(0).Top
        'PicDigit(intLoop).BorderStyle = 1
        PicDigit(intLoop).Width = PicDigit(0).Width
        PicDigit(intLoop).Height = PicDigit(0).Height
        If m_DigitSize = 0 Then
            PicDigit(intLoop).Left = PicDigit(intLoop - 1).Left + smImageWidth
        Else
            PicDigit(intLoop).Left = PicDigit(intLoop - 1).Left + lgImageWidth
        End If
    Next
End Function


Private Sub RepaintCtl() ' the main paint-routine
    Dim bdrFlags As Long, RT As RECT
    Dim intLoop As Integer
    Dim intBase As Integer
    
    UserControl.Cls
    
    Select Case m_BorderStyle
        Case 0, 4
        Case 1: bdrFlags = EDGE_ETCHED
        Case 2
            If m_Appearance = Thin Then bdrFlags = BDR_RAISEDINNER Else bdrFlags = BDR_RAISED
        Case 3
            If m_Appearance = Thin Then bdrFlags = BDR_SUNKENOUTER Else bdrFlags = BDR_SUNKEN
    End Select
    
    RT.Left = 0: RT.Right = ScaleWidth: RT.Top = 0: RT.Bottom = ScaleHeight
    DrawEdge UserControl.hdc, RT, bdrFlags, BF_RECT
    
    If m_BorderStyle = 4 Then
        UserControl.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), m_BorderColor, B
    End If
    
    If m_DigitSize = Small Then
        intBase = 110
    Else
        intBase = 210
    End If
    
    If m_FillChar = Blank Then
        intBase = intBase + 11
    End If
    
    For intLoop = 0 To PicDigit.UBound
        PicDigit(intLoop).Picture = LoadResPicture(intBase, vbResBitmap)
    Next
End Sub


Public Function UpdateValue()
    Dim strBuffer As String
    Dim strValue As String
    Dim intCurDigit As Integer
    Dim intBase As Integer
    Dim strChar As String
    Dim intChar As Integer
    Dim intOffset As Integer
    Dim intLoop As Integer
    Dim strTemp As String
    Dim strFmt As String
    
    strValue = m_Value
    strTemp = ""
    
    On Error GoTo UValue_err
    
    If m_DigitSize = Small Then
        intBase = 110
    Else
        intBase = 210
    End If
    
    For intLoop = 0 To m_DigitCount - 1
        PicDigit(intLoop).Picture = LoadResPicture(intBase, vbResBitmap)
    Next
    
    intCurDigit = 0
    intOffset = 0
    'strValue = ""
    Do While intCurDigit < Len(strValue) ' - 1
        strChar = Mid(strValue, intCurDigit + 1 + intOffset, 1)
        If strChar = "-" Then
            intChar = intBase + 14
            intCurDigit = intCurDigit + 1
            strBuffer = strBuffer & Chr(intChar)
        ElseIf strChar = "." Then
            intChar = Val(Asc(Mid(strBuffer, Len(strBuffer), 1))) - 10
            If Len(strBuffer) > 1 Then
                Mid(strBuffer, Len(strBuffer), 1) = Chr(intChar)
            Else
                Mid(strBuffer, 1, 1) = Chr(intChar)
            End If
            intOffset = 1
        Else
            intChar = intBase + Val(strChar)
            intCurDigit = intCurDigit + 1
            strBuffer = strBuffer & Chr(intChar)
        End If
    Loop
    
    Do While Len(strFmt) < (m_DigitCount - Len(strBuffer))
        If m_FillChar = Blank Then
            strFmt = strFmt & Chr(intBase + 11) 'blank
        Else
            strFmt = strFmt & Chr(intBase)  'zero
        End If
    Loop
    
    If Asc(Left(strBuffer, 1)) - intBase = 14 And m_FillChar = Zero Then
        strBuffer = Chr(intBase + 14) & strFmt & Mid(strBuffer, 2, Len(strBuffer))
    Else
        strBuffer = strFmt & strBuffer
    End If
    For intCurDigit = 0 To m_DigitCount - 1
'        If Mid(strBuffer, intCurDigit + 1, 1) <> "" Then
            intChar = Asc(Mid(strBuffer, intCurDigit + 1, 1))
            PicDigit(intCurDigit).Picture = LoadResPicture(intChar, vbResBitmap)
'        End If
    Next
Exit Function
UValue_err:
        For intLoop = 0 To PicDigit.UBound
            PicDigit(intLoop).Picture = LoadResPicture(intBase + 14, vbResBitmap)
        Next
        RaiseEvent Error
    Exit Function
End Function


Private Sub UserControl_AmbientChanged(PropertyName As String)
    RepaintCtl
    UpdateValue
End Sub

Private Sub UserControl_InitProperties()
    m_BorderStyle = Etched
    m_DigitSize = Small
    m_DigitCount = 4
    m_Value = 1234
    m_FillChar = Blank
    InitDisplay
    UpdateValue
    RepaintCtl
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Value = .ReadProperty("Value", 1234)
        m_BorderStyle = .ReadProperty("BStyle", 1)
        m_DigitSize = .ReadProperty("DSize", 1)
        m_DigitCount = .ReadProperty("DCount", 4)
        m_FillChar = .ReadProperty("LChar", 1)
    End With
    InitDisplay
    RepaintCtl
End Sub
Private Sub UserControl_Resize()
    UserControl.Width = (PicDigit(PicDigit.UBound).Left + PicDigit(PicDigit.UBound).Width + PicDigit(0).Left) * Screen.TwipsPerPixelX
    UserControl.Height = (PicDigit(PicDigit.UBound).Top + PicDigit(PicDigit.UBound).Height + PicDigit(0).Top) * Screen.TwipsPerPixelY
End Sub
Public Property Get DigitSize() As DigitSize
    DigitSize = m_DigitSize
End Property

Public Property Let DigitSize(ByVal vNewValue As DigitSize)
    m_DigitSize = vNewValue
    InitDisplay
    UserControl_Resize
    RepaintCtl
    UpdateValue
    PropertyChanged "DigitSize"
End Property


Public Property Get BorderStyle() As mhc_BorderStyle
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewValue As mhc_BorderStyle)
    m_BorderStyle = vNewValue
    RepaintCtl
    PropertyChanged "BorderStyle"
End Property

Public Property Get DigitCount() As Integer
    DigitCount = m_DigitCount
End Property

Public Property Let DigitCount(ByVal vNewValue As Integer)
    If vNewValue > 0 Then
        m_DigitCount = vNewValue
        InitDisplay
        UserControl_Resize
        RepaintCtl
        PropertyChanged "DigitCount"
    End If
End Property

Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Let Value(ByVal vNewValue As Double)
    m_Value = vNewValue
    UpdateValue
    PropertyChanged "Value"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Value", m_Value, 1234
        .WriteProperty "BStyle", m_BorderStyle, 1
        .WriteProperty "DSize", m_DigitSize, 1
        .WriteProperty "DCount", m_DigitCount, 4
        .WriteProperty "LChar", m_FillChar, 1
    End With
End Sub

Public Property Get LeadingChar() As LeadChar
    LeadingChar = m_FillChar
End Property

Public Property Let LeadingChar(ByVal vNewValue As LeadChar)
    m_FillChar = vNewValue
    RepaintCtl
    UpdateValue
    PropertyChanged "LeadingChar"
End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
    If Not frmAbout.Visible = True Then
        Load frmAbout
        frmAbout.Show 1
    End If
End Sub
