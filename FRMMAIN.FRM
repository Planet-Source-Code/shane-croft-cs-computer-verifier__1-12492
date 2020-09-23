VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   Caption         =   "CS Computer Verifier - Windows NT-2000"
   ClientHeight    =   7830
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10605
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Remove Computer from list"
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   4080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   4080
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   120
      TabIndex        =   37
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   6720
      TabIndex        =   35
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   8040
      TabIndex        =   34
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stored Information"
      Height          =   2295
      Left            =   2760
      TabIndex        =   13
      Top             =   2400
      Width           =   7815
      Begin VB.TextBox Text6 
         Height          =   855
         Index           =   1
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   975
         Index           =   1
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         Height          =   525
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   525
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   735
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Adapter"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   31
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Drive Info"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Memory"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Processer"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "OS"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Information Found"
      Height          =   2295
      Left            =   2760
      TabIndex        =   12
      Top             =   0
      Width           =   7815
      Begin VB.TextBox Text6 
         Height          =   855
         Index           =   0
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   975
         Index           =   0
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         Height          =   525
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   525
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   735
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Adapter"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   21
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Drive Info"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Memory"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Processer"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "OS"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Alerts"
      Height          =   255
      Left            =   9360
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   7200
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Verify Information on All"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Verify Information"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7455
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13070
            Text            =   "Status:"
            TextSave        =   "Status:"
            Object.ToolTipText     =   "Current Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:09 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "9/18/2000"
            Object.ToolTipText     =   "Current Date"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   0
      TabIndex        =   36
      Top             =   5160
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Alert Date"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   14711
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "50%"
      Height          =   255
      Left            =   5115
      TabIndex        =   10
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "100%"
      Height          =   255
      Index           =   0
      Left            =   10200
      TabIndex        =   9
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "0%"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Alerts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   10455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Computers:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Computer List"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.Menu menufile 
      Caption         =   "&File"
      Begin VB.Menu menureload 
         Caption         =   "Reload Computer List"
         Shortcut        =   {F1}
      End
      Begin VB.Menu menuline 
         Caption         =   "-"
      End
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu menusettings 
      Caption         =   "&Settings"
      Begin VB.Menu menuselect 
         Caption         =   "Select Computers"
         Shortcut        =   {F2}
      End
      Begin VB.Menu menuping 
         Caption         =   "Check computer Offline/Online Status"
         Shortcut        =   {F3}
      End
      Begin VB.Menu menuviewinfo 
         Caption         =   "View/Delete Saved Computer Information"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "&Help"
      Begin VB.Menu menuabout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu menuhelppage 
         Caption         =   "Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu menuweb 
         Caption         =   "Web Page"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private db As Database
Private rs As Recordset
Dim WithEvents sink As SWbemSink
Attribute sink.VB_VarHelpID = -1
Public Sub CompareData()
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
If Text2(1).Text = "" Then
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - There is nothing stored in the database for Operating System Information"
Else
If Text2(0).Text = Text2(1).Text Then
Else
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - Operating System information doesn't match the stored information." & ENTER & ENTER & "Stored information is:" & ENTER & Text2(1).Text & ENTER & ENTER & " Information found was:" & ENTER & Text2(0).Text
End If
End If

If text3(1).Text = "" Then
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - There is nothing stored in the database for Processor Information"
Else
If text3(0).Text = text3(1).Text Then
Else
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - Processor information doesn't match the stored information." & ENTER & ENTER & "Stored information is:" & ENTER & text3(1).Text & ENTER & ENTER & " Information found was:" & ENTER & text3(0).Text
End If
End If

If text4(1).Text = "" Then
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - There is nothing stored in the database for Memory Information"
Else
If text4(0).Text = text4(1).Text Then
Else
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - Memory information doesn't match the stored information." & ENTER & ENTER & "Stored information is:" & ENTER & text4(1).Text & ENTER & ENTER & " Information found was:" & ENTER & text4(0).Text
End If
End If

If text5(1).Text = "" Then
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - There is nothing stored in the database for Drive Information"
Else
If text5(0).Text = text5(1).Text Then
Else
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - Drive information doesn't match the stored information." & ENTER & ENTER & "Stored information is:" & ENTER & text5(1).Text & ENTER & ENTER & " Information found was:" & ENTER & text5(0).Text
End If
End If

If text6(1).Text = "" Then
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - There is nothing stored in the database for Adapter Information"
Else
If text6(0).Text = text6(1).Text Then
Else
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = "Computer: " & List1.Text & " - Adapter information doesn't match the stored information." & ENTER & ENTER & "Stored information is:" & ENTER & text6(1).Text & ENTER & ENTER & " Information found was:" & ENTER & text6(0).Text
End If
End If
End Sub

Public Sub reloadlist()
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
Private Sub Command1_Click()
On Error Resume Next
Set db = OpenDatabase(App.Path & "\Computers.mdb")
Set rs = db.OpenRecordset("SELECT * FROM Computer_Info " & "ORDER BY [Computer Name]")
frmmain.StatusBar1.Panels(1).Text = "Status: Please wait loading stored information on computer: " & List1.Text

Text2(0).Text = ""
text3(0).Text = ""
text4(0).Text = ""
text5(0).Text = ""
text6(0).Text = ""
Text2(1).Text = ""
text3(1).Text = ""
text4(1).Text = ""
text5(1).Text = ""
text6(1).Text = ""
Text7.Text = "0"
Text8.Text = ""
    
    Frame1.Caption = "Current Information Found on " & List1.Text
    Frame2.Caption = "Stored Information on " & List1.Text
    rs.FindFirst "[ID] = " & List1.ItemData(List1.ListIndex)
    
    Text2(1).Text = rs.Fields("Os") & ""
    text3(1).Text = rs.Fields("Processor Info") & ""
    text4(1).Text = rs.Fields("Memory") & ""
    text5(1).Text = rs.Fields("Drives") & ""
    text6(1).Text = rs.Fields("Adapter Info") & ""
    DoEvents
frmmain.StatusBar1.Panels(1).Text = "Status: Please wait receiving information from computer: " & List1.Text
Call GetAllData2
DoEvents
frmmain.StatusBar1.Panels(1).Text = "Status: Please wait comparing information found with saved information..."
Call CompareData
frmmain.StatusBar1.Panels(1).Text = "Status:"

End Sub

Private Sub Command2_Click()
On Error Resume Next
List2.Clear
ProgressBar1.Value = 0
List1.ListIndex = 0
Timer3.Enabled = True
ProgressBar1.Max = List1.ListCount
End Sub

Private Sub Command3_Click()
On Error Resume Next
ListView1.ListItems.Clear
End Sub



Private Sub Command4_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
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

ProgressBar1.Max = List1.ListCount
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
    Dim f As frmreadalert
    Set f = New frmreadalert
    f.Text2.Text = ListView1.SelectedItem
f.Text1.Text = ListView1.SelectedItem.SubItems(1)
f.Show
End Sub

Private Sub menuabout_Click()
On Error Resume Next
frmAbout.Show
End Sub

Private Sub menuexit_Click()
End
End Sub

Private Sub menuhelppage_Click()
On Error Resume Next
Call ShellExecute(hwnd, "Open", App.Path & "\Help\Help.htm", "", App.Path, 1)
End Sub

Private Sub menuping_Click()
frmping.Show
End Sub

Private Sub menureload_Click()
On Error Resume Next
Call reloadlist
End Sub

Private Sub menuselect_Click()
On Error Resume Next
frmcomputers.Show
End Sub

Private Sub menuviewinfo_Click()
On Error Resume Next
frmdatabase.Show
End Sub

Private Sub menuweb_Click()
On Error Resume Next
Call ShellExecute(hwnd, "Open", "http://www.croftssoftware.com", "", App.Path, 1)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label2(0).Caption = "Total Computers: " & List1.ListCount
If List1.ListCount = 0 Then
Command1.Enabled = False
Command2.Enabled = False
Else
Command1.Enabled = True
Command2.Enabled = True
End If
If List1.ListCount = 0 Then
menuviewinfo.Enabled = False
menuping.Enabled = False
Else
menuviewinfo.Enabled = True
menuping.Enabled = True
End If
End Sub
Private Sub sink_OnObjectReady(ByVal objWbemObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
Dim i As Integer
i = Text7.Text
Set adapter = GetObject("winmgmts:Win32_NetworkAdapterConfiguration=" & i & "")
Description = adapter.Description

text6(0).Text = text6(0).Text & Description & ENTER

If IsNull(adapter.MACAddress) Then
    text6(0).Text = text6(0).Text & "No MAC Address" & ENTER
    text6(0).Text = text6(0).Text & "" & ENTER
Else
    text6(0).Text = text6(0).Text & "Mac: " & adapter.MACAddress & ENTER
    text6(0).Text = text6(0).Text & "" & ENTER
End If
 

Text7.Text = i + 1
End Sub

Public Sub GetAllData2()
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
'Os Information
Set SystemSet = GetObject("winmgmts:\\" & List1.Text).InstancesOf("Win32_OperatingSystem")

For Each System In SystemSet
    Text2(0).Text = Text2(0).Text & System.Caption & ENTER
    Text2(0).Text = Text2(0).Text & System.Manufacturer & ENTER
    Text2(0).Text = Text2(0).Text & System.BuildType & ENTER
    Text2(0).Text = Text2(0).Text & " Version: " + System.Version & ENTER
    Text2(0).Text = Text2(0).Text & " Serial Number: " + System.SerialNumber & ENTER
Next

'Adapter information
    ' Create a sink to recieve the results of the enumeration
    Set sink = New SWbemSink
    
    ' Connect to root\cimv2.
    Set adapter = GetObject("winmgmts:\\" & List1.Text)
' Perform the asynchronous enumeration of processes
adapter.InstancesOfAsync sink, "Win32_NetworkAdapter"

'Processor Information
Set obj = GetObject("winmgmts:\\" & List1.Text).InstancesOf("Win32_Processor")


            For Each obj2 In obj
            text3(0).Text = text3(0).Text & obj2.Caption & ENTER
            text3(0).Text = text3(0).Text & "Speed: " & obj2.currentclockspeed & " Mhz" & ENTER

Next

'get memory
Set obj = GetObject("winmgmts:\\" & List1.Text).InstancesOf("Win32_PhysicalMemory")
Dim i As String

            For Each obj2 In obj
            Text8.Text = obj2.capacity
            i = Text8.Text
            ii = i / 1024
            iii = ii / 1024
            text4(0).Text = text4(0).Text & iii & " MB" & " Chip" & ENTER
Next

'Drive Info
On Error GoTo driveerror
Set obj = GetObject("winmgmts:\\" & List1.Text).InstancesOf("Win32_DiskDrive")

            For Each obj2 In obj
            Text8.Text = obj2.Size
            i = Text8.Text
            ii = i / 1024
            iii = ii / 1024
            iiii = iii / 1024
           text5(0).Text = text5(0).Text & obj2.Caption & " - " & Left$(iiii, 5) & " GB" & ENTER
          
Next
  Exit Sub
driveerror:
  text5(0).Text = text5(0).Text & "Removable Drive"
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If List2.ListCount = List1.ListCount Then
Timer2.Enabled = False
Else
List1.ListIndex = List1.ListIndex + 1
Timer3.Enabled = True
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Set db = OpenDatabase(App.Path & "\Computers.mdb")
Set rs = db.OpenRecordset("SELECT * FROM Computer_Info " & "ORDER BY [Computer Name]")
frmmain.StatusBar1.Panels(1).Text = "Status: Please wait loading stored information on computer: " & List1.Text

Text2(0).Text = ""
text3(0).Text = ""
text4(0).Text = ""
text5(0).Text = ""
text6(0).Text = ""
Text2(1).Text = ""
text3(1).Text = ""
text4(1).Text = ""
text5(1).Text = ""
text6(1).Text = ""
Text7.Text = "0"
Text8.Text = ""
    
    Frame1.Caption = "Current Information Found on " & List1.Text
    Frame2.Caption = "Stored Information on " & List1.Text
    rs.FindFirst "[ID] = " & List1.ItemData(List1.ListIndex)
    
    Text2(1).Text = rs.Fields("Os") & ""
    text3(1).Text = rs.Fields("Processor Info") & ""
    text4(1).Text = rs.Fields("Memory") & ""
    text5(1).Text = rs.Fields("Drives") & ""
    text6(1).Text = rs.Fields("Adapter Info") & ""
    DoEvents
frmmain.StatusBar1.Panels(1).Text = "Status: Please wait receiving information from computer: " & List1.Text
Call GetAllData2
DoEvents
frmmain.StatusBar1.Panels(1).Text = "Status: Please wait comparing information found with saved information..."
Call CompareData
frmmain.StatusBar1.Panels(1).Text = "Status:"
List2.AddItem List1.Text
ProgressBar1.Value = ProgressBar1.Value + 1
Timer3.Enabled = False
Timer2.Enabled = True
End Sub
