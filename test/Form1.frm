VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7635
   StartUpPosition =   3  'Padrão Windows
   Begin VB.PictureBox Picture1 
      Height          =   3570
      Left            =   3960
      ScaleHeight     =   3510
      ScaleWidth      =   3510
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   3570
   End
   Begin VB.TextBox Text1 
      Height          =   348
      Left            =   252
      TabIndex        =   2
      Text            =   "http://www.vbforums.com"
      Top             =   168
      Width           =   5490
   End
   Begin VB.CommandButton cmdCommand1 
      Height          =   480
      Left            =   6960
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Square"
      Height          =   192
      Left            =   5940
      TabIndex        =   0
      Top             =   252
      Value           =   1  'Marcado
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   3570
      Left            =   252
      Stretch         =   -1  'True
      Top             =   750
      Width           =   3570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' QR Code generator library (VB6/VBA)
'
' Copyright (c) Project Nayuki. (MIT License)
' https://www.nayuki.io/page/qr-code-generator-library
'
' Copyright (c) wqweto@gmail.com (MIT License)
'
'=========================================================================
Option Explicit
DefObj A-Z

Private Sub cmdCommand1_Click()
Picture1.Picture = Image1.Picture
SavePicture Picture1, App.Path & "\" & Text1.Text & ".jpg"
MsgBox "Salvo em:" & vbNewLine & App.Path & "\" & Text1.Text & ".jpg", vbExclamation, "Imagem salva:"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim baBarCode()     As Byte
    Dim lQrSize         As Long
    Dim lModuleSize     As Long
    
    If KeyCode = 67 And Shift = vbCtrlMask Then
        Clipboard.Clear
        Clipboard.SetData Image1.Picture
    ElseIf KeyCode = 67 And Shift = (vbCtrlMask Or vbShiftMask) Then
        If QRCodegenEncode(Text1.Text, baBarCode) Then
            lQrSize = QRCodegenGetSize(baBarCode)
            lModuleSize = Int((Image1.Width * 15) / (lQrSize * Screen.TwipsPerPixelX) + 0.5)
            Clipboard.Clear
            Clipboard.SetData QRCodegenConvertToPicture(baBarCode, vbRed, ModuleSize:=lModuleSize, SquareModules:=(Check1.Value = vbChecked))
        End If
    ElseIf KeyCode = 67 And Shift = vbAltMask Then
        If QRCodegenEncode(Text1.Text, baBarCode) Then
            lQrSize = QRCodegenGetSize(baBarCode)
            lModuleSize = Int((Image1.Width * 15) / (lQrSize * Screen.TwipsPerPixelX) + 0.5)
            Clipboard.Clear
            Clipboard.SetData QRCodegenResizePicture(QRCodegenResizePicture(QRCodegenConvertToPicture(baBarCode, vbBlue, ModuleSize:=lModuleSize, SquareModules:=(Check1.Value = vbChecked)), 2000, 2000), 500, 500)
        End If
    End If
End Sub

Private Sub Form_Load()
    Text1_Change
End Sub

Private Sub Form_Resize()
    Dim lWidth          As Long
    Dim lHeight         As Long
    
    If WindowState <> vbMinimized Then
        lWidth = ScaleWidth - Image1.Left - Image1.Left
        lHeight = ScaleHeight - Image1.Top - Image1.Left
        If lWidth > lHeight Then
            lWidth = lHeight
        End If
        Image1.Width = lWidth
        Image1.Height = lWidth
    End If
End Sub

Private Sub Text1_Change()
    Set Image1.Picture = QRCodegenBarcode(Text1.Text, SquareModules:=(Check1.Value = vbChecked))
End Sub

Private Sub Check1_Click()
    Text1_Change
End Sub
