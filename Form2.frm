VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3720
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox txtHex 
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      RightMargin     =   1
      TextRTF         =   $"Form2.frx":0000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    End
End Sub

Private Sub Form_Load()
    txtHex.Width = Screen.Width - (60 * 2)
    txtHex.Height = Screen.Height - (60 * 2)
    ' put the rich textbox in a nice and big possision on screen!
End Sub
