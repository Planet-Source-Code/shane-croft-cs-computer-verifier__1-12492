VERSION 5.00
Begin VB.Form frmrecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information on "
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmrecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   1335
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3960
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      Height          =   1215
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   525
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "Adapter Information:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Drive Information:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Memory:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Processer Information:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Operating System Information:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmrecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

