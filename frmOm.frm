VERSION 5.00
Begin VB.Form frmOm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Om SuperWriter 1.0"
   ClientHeight    =   2130
   ClientLeft      =   6150
   ClientTop       =   4050
   ClientWidth     =   3570
   Icon            =   "frmOm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3570
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3000
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "  Klockan är:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Idag är det den:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Ett underbart program som utvecklats av Joakim Bengtsson"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmOm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Date$
Label5.Caption = Time$
End Sub
