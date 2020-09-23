VERSION 5.00
Object = "*\APrjLCD.vbp"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1140
   ClientLeft      =   1290
   ClientTop       =   1395
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   6420
   Begin Project2.LCDDisplay LCDDisplay1 
      Height          =   675
      Left            =   1200
      TabIndex        =   1
      Top             =   180
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1191
      DigitCount      =   6
      DigitSize       =   1
      DispValue       =   90000
   End
   Begin FLWCtrls.FWDial FWDial1 
      Height          =   915
      Left            =   4380
      TabIndex        =   0
      Top             =   120
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      Max             =   9999
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl11_GotFocus()

End Sub

Private Sub Form_Load()
    'LCDDisplay1.DigitSize = LCD_Large
End Sub

Private Sub FWDial1_Change()
    LCDDisplay1.Value = FWDial1.Value
End Sub
