VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2520
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "" Then If MsgBox("harap isi data", vbOKOnly + vbCritical, "Harap isi") = vbOK Then Exit Sub
Combo1.AddItem Combo1.Text
Combo1.Text = ""
Combo1.SetFocus
End Sub

Private Sub Form_Load()
Form1.Caption = "Combo nama-nama buah"
Command1.Caption = "Tambah Item"
Combo1.Text = "item Awal"
End Sub
