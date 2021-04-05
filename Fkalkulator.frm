VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command20 
      Caption         =   "Command20"
      Height          =   495
      Left            =   6120
      TabIndex        =   20
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Command19"
      Height          =   495
      Left            =   4680
      TabIndex        =   18
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Command18"
      Height          =   495
      Left            =   6120
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Command15"
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   495
      Left            =   6120
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Angka10 
      Caption         =   "10"
      Height          =   495
      Index           =   10
      Left            =   1800
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Angka9 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Angka8 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   1800
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Angka7 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Angka6 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   3240
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Angka5 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Angka4 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Angka3 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Angka2 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Angka1 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim operator As String
Dim hapuslayar As Boolean
Dim operasi1 As Double, operasi2 As Double


Private Sub Angka1_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka1.Item(1).Caption
End Sub

Private Sub Angka10_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka10.Item(10).Caption
End Sub

Private Sub Angka2_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka2.Item(2).Caption
End Sub

Private Sub Angka3_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka3.Item(3).Caption
End Sub

Private Sub Angka4_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka4.Item(4).Caption
End Sub

Private Sub Angka5_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka5.Item(5).Caption
End Sub

Private Sub Angka6_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka6.Item(6).Caption
End Sub

Private Sub Angka7_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka7.Item(7).Caption
End Sub

Private Sub Angka8_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka8.Item(8).Caption
End Sub

Private Sub Angka9_Click(Index As Integer)
Label1.Caption = Label1.Caption + Angka9.Item(9).Caption
End Sub

Private Sub Command10_Click()
Label1.Caption = ""
End Sub

Private Sub Command12_Click()
If hapuslayar Then
    Label1.Caption = ""
    hapuslayar = False
End If
If InStr(Label1.Caption, ".") Then
    Exit Sub
Else
Label1.Caption = Label1.Caption + "."
End If


End Sub

Private Sub Command13_Click()
Label1.Caption = -Val(Label1.Caption)
End Sub

Private Sub Command14_Click()
If Val(Label1.Caption) <> 0 Then Label1.Caption = 1 / Val(Label1.Caption)
End Sub

Private Sub Command15_Click()
operasi1 = Val(Label1.Caption)
operator = "+"
Label1.Caption = ""
End Sub

Private Sub Command16_Click()
operasi1 = Val(Label1.Caption)
operator = "-"
Label1.Caption = ""
End Sub

Private Sub Command17_Click()
operasi1 = Val(Label1.Caption)
operator = "*"
Label1.Caption = ""
End Sub

Private Sub Command18_Click()
operasi1 = Val(Label1.Caption)
operator = "/"
Label1.Caption = ""
End Sub

Private Sub Command19_Click()
Dim hasil As Double
operasi2 = Val(Label1.Caption)
If operator = "+" Then hasil = operasi1 + operasi2
If operator = "-" Then hasil = operasi1 - operasi2
If operator = "*" Then hasil = operasi1 * operasi2
If operator = "/" And operasi2 <> 0 Then hasil = operasi1 / operasi2
Label1.Caption = hasil
operator = ""
hapuslayar = True

End Sub

Private Sub Command20_Click()
If MsgBox("Yakin ingin keluar ? ", vbYesNo + vbCritical, "Keluar") = vbYes Then Unload Me
End Sub

Private Sub Form_Load()
Label1.FontSize = 20
Form1.Caption = "Kalkulator"
Label1.Alignment = 1
Label1.Caption = ""
Command10.Caption = "C"
Command12.Caption = "."
Command13.Caption = "+/-"
Command14.Caption = "1/X"
Command15.Caption = "+"
Command16.Caption = "-"
Command17.Caption = "*"
Command18.Caption = "/"
Command19.Caption = "="
Command20.Caption = "Keluar"
End Sub
