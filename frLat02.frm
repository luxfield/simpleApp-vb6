VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   3735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
      Begin VB.OptionButton Option8 
         Caption         =   "Option1"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Option1"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option1"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option1"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If MsgBox("Yakin Ingin keluar ?", vbYesNo + vbCritical, "Keluar") = vbYes Then
 Unload Form1
 End If
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = vbYellow
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = vbBlue
End Sub
Private Sub Command1_prop()
Command1.Caption = "EXIT"
Command1.FontBold = True
End Sub
Private Sub Form_Load()
Call Label1_prop
Call Label2_prop
Call Label3_prop
Call Text1_prop
Call Text2_prop
Call Text3_prop
Call Frame_prop
Call Option1_prop
Call Option2_prop
Call Option3_prop
Call Option4_prop
Call Option5_prop
Call Option6_prop
Call Option7_prop
Call Option8_prop
Call Command1_prop
End Sub
Private Sub Frame_prop()
With Frame1
    .Caption = "Operator"
    .FontSize = 8
    .FontItalic = True
    .FontBold = True
End With
End Sub
Private Sub Label1_prop()
Label1.Caption = "Bilangan 1"
Label1.FontBold = True
Label1.FontSize = 8
End Sub
Private Sub Label2_prop()
Label2.Caption = "Bilangan 2"
Label2.FontBold = True
Label2.FontSize = 8
End Sub
Private Sub Label3_prop()
Label3.Caption = "Bilangan 3"
Label3.FontBold = True
Label3.FontSize = 8
End Sub
Private Sub Text1_prop()
Text1.Text = ""
Text1.FontSize = 10
Text1.FontBold = True
End Sub
Private Sub Text2_prop()
Text2.Text = ""
Text2.FontSize = 10
Text2.FontBold = True
End Sub
Private Sub Text3_prop()
Text3.Text = ""
Text3.FontSize = 10
Text3.FontBold = True
Text3.Enabled = False
End Sub
Private Sub Option1_prop()
Option1.Caption = "+"
Option1.FontBold = True
Option1.FontSize = 8
End Sub
Private Sub Option2_prop()
Option2.Caption = "-"
Option2.FontBold = True
Option2.FontSize = 8
End Sub
Private Sub Option3_prop()
Option3.Caption = "*"
Option3.FontBold = True
Option3.FontSize = 8
End Sub
Private Sub Option4_prop()
Option4.Caption = "/"
Option4.FontBold = True
Option4.FontSize = 8
End Sub
Private Sub Option5_prop()
Option5.Caption = "\"
Option5.FontBold = True
Option5.FontSize = 8
End Sub
Private Sub Option6_prop()
Option6.Caption = "^"
Option6.FontBold = True
Option6.FontSize = 8
End Sub
Private Sub Option7_prop()
Option7.Caption = "&&"
Option7.FontBold = True
Option7.FontSize = 8
End Sub
Private Sub Option8_prop()
Option8.Caption = "Mod"
Option8.FontBold = True
Option8.FontSize = 8
End Sub
Private Sub Option1_Click()
If Text1.Text <> "" Or Text2.Text <> "" Then
Text3.Text = CInt(Text1.Text) + CInt(Text2.Text)
End If
End Sub

Private Sub Option2_Click()
If Text1.Text <> "" Or Text2.Text <> "" Then
Text3.Text = CInt(Text1.Text) - CInt(Text2.Text)
End If
End Sub

Private Sub Option3_Click()
If Text1.Text <> "" Or Text2.Text <> "" Then
Text3.Text = CInt(Text1.Text) * CInt(Text2.Text)
End If
End Sub

Private Sub Option4_Click()
If Text1.Text <> "" Or Text2.Text <> "" Then
Text3.Text = CInt(Text1.Text) / CInt(Text2.Text)
End If
End Sub

Private Sub Option5_Click()
If Text1.Text <> "" Or Text2.Text <> "" Then
Text3.Text = CInt(Text1.Text) \ CInt(Text2.Text)
End If
End Sub

Private Sub Option6_Click()
If Text1.Text <> "" Or Text2.Text <> "" Then
Text3.Text = CInt(Text1.Text) ^ CInt(Text2.Text)
End If
End Sub

Private Sub Option7_Click()
If Text1.Text <> "" Or Text2.Text <> "" Then
Text3.Text = CInt(Text1.Text) And CInt(Text2.Text)
End If
End Sub

Private Sub Option8_Click()
If Text1.Text <> "" Or Text2.Text <> "" Then
Text3.Text = CInt(Text1.Text) Mod CInt(Text2.Text)
End If
End Sub
