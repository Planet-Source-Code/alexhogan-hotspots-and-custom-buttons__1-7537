VERSION 5.00
Begin VB.Form frmExplain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Explaination"
   ClientHeight    =   4830
   ClientLeft      =   2925
   ClientTop       =   2445
   ClientWidth     =   3720
   Icon            =   "frmExplain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtExplain 
      Height          =   3975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmExplain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Info As String

Open "Info.txt" For Input As #1
txtExplain.Text = Input(LOF(1), 1)

End Sub

