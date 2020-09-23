VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MD5 Hash"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Hash"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Output"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   645
      Width           =   735
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Input"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   165
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MD5 As clsMD5

Private Sub Cmd_Click()
    Set MD5 = New clsMD5
    Text2 = MD5.CalculateMD5(Text1)
End Sub
