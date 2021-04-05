VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text1.Alignment = 2
Else
    If Check1.Value = 0 Then
        Text1.Alignment = 0
    End If
End If
End Sub

Private Sub Command1_Click()
Text1.Text = "Latihan Program Text"
End Sub

Private Sub Command2_Click()
Dim tgl
'format tampilan tanggal
tgl = Format(Date, "dddd, d mmmm yyyy")
Text1.Text = "Sekarang hari dan tanggal : " + tgl
End Sub

Private Sub Command3_Click()
Text1.Text = ""
End Sub


Private Sub Form_Load()
Command1.Caption = "judul"
Command2.Caption = "tampil tanggal"
Command3.Caption = "hapus layar"
Check1.Caption = "Text rata tengah"
Text1.FontSize = 16
End Sub
