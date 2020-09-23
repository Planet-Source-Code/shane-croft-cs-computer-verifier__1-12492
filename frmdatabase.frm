VERSION 5.00
Begin VB.Form frmdatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Computers"
   ClientHeight    =   6045
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   3390
   Icon            =   "frmdatabase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   3390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Remove All Computers From Database"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Computer From Database"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1088
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   128
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Double click to view information"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Computers in Database:"
      Height          =   255
      Left            =   135
      TabIndex        =   2
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select Computer"
      Height          =   255
      Left            =   135
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmdatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Database
Dim rs As Recordset
Public Sub deletedatabase()
On Error Resume Next
List1.ListIndex = 0
Do Until List1.ListCount = 0
rs.FindFirst "[ID] = " & List1.ItemData(List1.ListIndex)
rs.Delete
List1.RemoveItem List1.ListIndex
Loop

End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
rs.FindFirst "[ID] = " & List1.ItemData(List1.ListIndex)
rs.Delete
List1.RemoveItem List1.ListIndex
frmmain.List1.Clear
Call frmmain.reloadlist
End Sub

Private Sub Command3_Click()
On Error Resume Next
frmconfirm.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
    Set db = OpenDatabase(App.Path & "\Computers.mdb")
    Set rs = db.OpenRecordset("SELECT * FROM Computer_Info " & "ORDER BY [Computer Name]")
    
    ' Populate the list box
    Do Until rs.EOF
        List1.AddItem rs.Fields("Computer Name")
        List1.ItemData(List1.NewIndex) = rs.Fields("ID")
        
        rs.MoveNext
        
    Loop

End Sub
Private Sub List1_DblClick()
On Error Resume Next
    Dim f As frmrecord
    Set f = New frmrecord
    frmmain.StatusBar1.Panels(1).Text = "Status: Please wait loading stored information on computer: " & List1.Text
    rs.FindFirst "[ID] = " & List1.ItemData(List1.ListIndex)
    
    f.Text1.Text = rs.Fields("Os") & ""
    f.Text2.Text = rs.Fields("Processor Info") & ""
    f.Text3.Text = rs.Fields("Memory") & ""
    f.Text4.Text = rs.Fields("Drives") & ""
    f.Text5.Text = rs.Fields("Adapter Info") & ""
    
    f.Show
    f.Caption = "Information on " & rs.Fields("Computer Name")
frmmain.StatusBar1.Panels(1).Text = "Status:"

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label2.Caption = "Total Computers in Database: " & List1.ListCount
End Sub
