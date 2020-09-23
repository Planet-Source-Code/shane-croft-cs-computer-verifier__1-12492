VERSION 5.00
Begin VB.Form frmmsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computers not added to database"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmmsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   6375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6000
      Top             =   3240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Computers:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Computers not added to database"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmmsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub List_Add(List As ListBox, txt As String)
On Error Resume Next
    List1.AddItem txt
End Sub

Public Sub List_Load(thelist As ListBox, FileName As String)
    'Loads a file to a list box
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        Call List_Add(List1, TheContents$)
    Loop Until EOF(fFile)
    Close fFile
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Call List_Load(List1, App.Path & "\Errors.ini")
DoEvents
List1.ListIndex = 0
If List1.Text = "" Then
List1.RemoveItem List1.ListIndex
Else
End If

End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Total Computers: " & List1.ListCount
End Sub
