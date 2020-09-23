VERSION 5.00
Begin VB.Form frmcomputers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Computers"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   Icon            =   "frmcomputers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4920
      Top             =   3360
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Add to list"
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ok"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4920
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Computer List"
      Height          =   3255
      Left            =   5280
      TabIndex        =   8
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command5 
         Caption         =   "Remove All"
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove"
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Height          =   2595
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Total Computers:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Load computers by Domain"
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton Command3 
         Caption         =   "Add -->"
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add All -->"
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get Computers"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Total Computers:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Computers"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Please Choose a Domain (NT or 2000 Only)"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Specify a computer name"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   3255
   End
End
Attribute VB_Name = "frmcomputers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub List_Add(List As ListBox, txt As String)
On Error Resume Next
    List2.AddItem txt
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
        Call List_Add(List2, TheContents$)
    Loop Until EOF(fFile)
    Close fFile
End Sub

Public Sub List_Save(thelist As ListBox, FileName As String)
    'Save a listbox as FileName
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List2.List(Save)
    Next Save
    Close fFile
End Sub

Private Sub Command1_Click()
On Error Resume Next
MousePointer = vbHourglass

Dim container As IADsContainer
Dim containername As String
containername = Combo1.Text

Set container = GetObject("WinNT://" & containername)

container.Filter = Array("Computer")
Dim computer As IADsComputer
For Each computer In container
List1.AddItem computer.Name
Next
DoEvents
MousePointer = 0
End Sub

Private Sub Command2_Click()
On Error Resume Next
Do Until List1.ListCount = 0 Or List2.ListCount = 10
If List2.ListCount < 10 Then
List1.ListIndex = 0
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Else
MsgBox "Sorry you can only have up to 10 computers"
End If
Loop
End Sub

Private Sub Command3_Click()
On Error Resume Next
If List2.ListCount < 10 Then
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Else
MsgBox "Sorry you can only have up to 10 computers"
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
List2.RemoveItem List2.ListIndex
End Sub

Private Sub Command5_Click()
On Error Resume Next
List2.Clear
End Sub

Private Sub Command6_Click()
On Error Resume Next
List2.ListIndex = 0
Call List_Save(List2, App.Path & "\Computer_List.ini")
If List2.ListCount = 0 Or List2.Text = "" Then
Else
frmstoredata.Show
DoEvents
End If
frmmain.List1.Clear
Unload Me
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
On Error Resume Next
If List2.ListCount < 10 Then
List2.AddItem Text1.Text
Text1.Text = ""
Else
MsgBox "Sorry you can only have up to 10 computers"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call List_Load(List2, App.Path & "\Computer_List.ini")
DoEvents
List2.ListIndex = 0
If List2.Text = "" Then
List2.RemoveItem List2.ListIndex
Else
End If

Combo1.AddItem frmmain.Winsock1.LocalHostName
Dim namespace As IADsContainer
Dim domain As IADs
 'Loads Combo box1 with all the current domains
Set namespace = GetObject("WinNT:")

For Each domain In namespace
Combo1.AddItem domain.Name
Next

End Sub

Private Sub List1_DblClick()
On Error Resume Next
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Text1.Text = "" Then
Exit Sub
End If
If KeyCode = vbKeyReturn Then
 Call Command8_Click
 DoEvents
 End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label3.Caption = "Total Computers: " & List1.ListCount
Label4.Caption = "Total Computers: " & List2.ListCount
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Text1.Text = "" Then
Command8.Enabled = False
Else
Command8.Enabled = True
End If

If List1.ListCount = 0 Then
Command2.Enabled = False
Command3.Enabled = False
Else
Command2.Enabled = True
Command3.Enabled = True
End If

If Combo1.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

If List2.ListCount = 0 Then
Command4.Enabled = False
Command5.Enabled = False
Else
Command4.Enabled = True
Command5.Enabled = True
End If
End Sub
