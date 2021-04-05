VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Label1.FontSize = 18
Else
    If Check1.Value = 0 Then Label1.FontSize = 12
    End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Label1.Alignment = 2
Else
    If Check2.Value = 0 Then Label1.Alignment = 0
    End If
End Sub
Private Sub Form_Load()
Check1.Enabled = False
Check2.Enabled = False
Check1.Caption = "Font Besar"
Check2.Caption = "Teks rata tengah"
Label1.Caption = ""
Form1.Caption = "List nama teman-teman"
List1.AddItem "Ferdy"
List1.AddItem "Irwan"
List1.AddItem "Anelka"
List1.AddItem "Fauzan"
List1.AddItem "Kinarya"
List1.AddItem "Penaldy Ponco"
List1.AddItem "Fadia"
List1.AddItem "Kahla"
List1.AddItem "Khadafi"
End Sub

Private Sub List1_Click()
Label1.Caption = List1.List(List1.ListIndex)
Check1.Enabled = True
Check2.Enabled = True
End Sub
