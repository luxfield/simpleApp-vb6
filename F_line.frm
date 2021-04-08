VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.CommandButton Hasil 
         Caption         =   "Hasil"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Gambar"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   840
         Y1              =   240
         Y2              =   1080
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
