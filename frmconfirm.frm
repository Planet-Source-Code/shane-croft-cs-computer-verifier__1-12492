VERSION 5.00
Begin VB.Form frmconfirm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "???"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2265
   Icon            =   "frmconfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   2265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Are you sure you want to remove all the computers from the list?"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmconfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Private Sub Command1_Click()
On Error Resume Next
frmdatabase.List1.ListIndex = 0
Do Until frmdatabase.List1.ListCount = 0
rs.FindFirst "[ID] = " & frmdatabase.List1.ItemData(frmdatabase.List1.ListIndex)
rs.Delete
frmdatabase.List1.RemoveItem frmdatabase.List1.ListIndex
Loop
frmmain.List1.Clear
Call frmmain.reloadlist
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    Set db = OpenDatabase(App.Path & "\Computers.mdb")
    Set rs = db.OpenRecordset("SELECT * FROM Computer_Info " & "ORDER BY [Computer Name]")

End Sub
