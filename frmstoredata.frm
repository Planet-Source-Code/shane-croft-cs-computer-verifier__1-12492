VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmstoredata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Storing Information on all selected Computers"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   Icon            =   "frmstoredata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List4 
      Height          =   450
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   4080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   4080
   End
   Begin VB.ListBox List3 
      Height          =   1035
      Left            =   6000
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check Online/Offline Status First. (Recommended)"
      Height          =   255
      Left            =   2025
      TabIndex        =   24
      Top             =   4800
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start"
      Height          =   375
      Left            =   3405
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2400
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3720
      TabIndex        =   21
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   3360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer timercount 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5385
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox text1 
      DataField       =   "Computer Name"
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   7
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox text2 
      DataField       =   "Os"
      Height          =   765
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   435
      Width           =   3375
   End
   Begin VB.TextBox text3 
      DataField       =   "Processor Info"
      Height          =   525
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox text4 
      DataField       =   "Memory"
      Height          =   525
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox text5 
      DataField       =   "Drives"
      Height          =   765
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox text6 
      DataField       =   "Adapter Info"
      Height          =   765
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3240
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Computers:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "0%"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "100%"
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   16
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "50%"
      Height          =   255
      Left            =   3825
      TabIndex        =   15
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Computer Name:"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   13
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Os:"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   12
      Top             =   435
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Processor Info:"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Memory:"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Drives:"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   9
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Adapter Info:"
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   8
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Computer List"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmstoredata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private db As Database
Private rs As Recordset
Dim WithEvents sink As SWbemSink
Attribute sink.VB_VarHelpID = -1
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
   hName As Long
   hAliases As Long
   hAddrType As Integer
   hLength As Integer
   hAddrList As Long
End Type

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADescription_Len) As Byte
   szSystemStatus(0 To WSASYS_Status_Len) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Public Sub List_Save(thelist As ListBox, FileName As String)
    'Save a listbox as FileName
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List3.List(Save)
    Next Save
    Close fFile
End Sub

Function HiByte(ByVal wParam As Integer)
   
   HiByte = wParam \ &H100 And &HFF&
   
End Function

Function LoByte(ByVal wParam As Integer)
   
   LoByte = wParam And &HFF&
   
End Function

Sub SocketsInitialize()
   
   Dim WSAD As WSADATA
   Dim iReturn As Integer
   Dim sLowByte As String, sHighByte As String, sMsg As String
   
   iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
   
   If iReturn <> 0 Then
      MsgBox "Winsock.dll is not responding."
      End
   End If
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      sHighByte = Trim$(Str$(HiByte(WSAD.wVersion)))
      sLowByte = Trim$(Str$(LoByte(WSAD.wVersion)))
      sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
      sMsg = sMsg & " is not supported by winsock.dll "
      MsgBox sMsg
      End
   End If
   
   If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
      sMsg = "This application requires a minimum of "
      sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
      MsgBox sMsg
      End
   End If
   
End Sub

Sub SocketsCleanup()
   Dim lReturn As Long
   
   lReturn = WSACleanup()
   
   If lReturn <> 0 Then
      MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
      End
   End If
   
End Sub
Public Sub GetAllData()
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
'Os Information
Set SystemSet = GetObject("winmgmts:\\" & List1.Text).InstancesOf("Win32_OperatingSystem")

For Each System In SystemSet
    text2.Text = text2.Text & System.Caption & ENTER
    text2.Text = text2.Text & System.Manufacturer & ENTER
    text2.Text = text2.Text & System.BuildType & ENTER
    text2.Text = text2.Text & " Version: " + System.Version & ENTER
    text2.Text = text2.Text & " Serial Number: " + System.SerialNumber & ENTER
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
            text3.Text = text3.Text & obj2.Caption & ENTER
            text3.Text = text3.Text & "Speed: " & obj2.currentclockspeed & " Mhz" & ENTER

Next

'get memory
Set obj = GetObject("winmgmts:\\" & List1.Text).InstancesOf("Win32_PhysicalMemory")
Dim i As String

            For Each obj2 In obj
            Text8.Text = obj2.capacity
            i = Text8.Text
            ii = i / 1024
            iii = ii / 1024
            text4.Text = text4.Text & iii & " MB" & " Chip" & ENTER
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
           text5.Text = text5.Text & obj2.Caption & " - " & Left$(iiii, 5) & " GB" & ENTER
          
Next
  Exit Sub
driveerror:
  text5.Text = text5.Text & "Removable Drive"

End Sub
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
Public Sub reloadlist()
On Error Resume Next
Call List_Load(List1, App.Path & "\Computer_List.ini")
DoEvents
List1.ListIndex = 0
If List1.Text = "" Then
List1.RemoveItem List1.ListIndex
Else
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
rs.AddNew
text1.Text = ""
text2.Text = ""
text3.Text = ""
text4.Text = ""
text5.Text = ""
text6.Text = ""
End Sub

Private Sub Command2_Click()
On Error Resume Next
   rs.Fields("Computer Name") = text1.Text
If text2.Text = "" Then
rs.Fields("Os") = " "
Else
rs.Fields("Os") = text2.Text & ""
End If

If text3.Text = "" Then
rs.Fields("Processor Info") = " "
Else
 rs.Fields("Processor Info") = text3.Text & ""
End If
If text4.Text = "" Then
rs.Fields("Memory") = " "
Else
 rs.Fields("Memory") = text4.Text & ""
End If
If text5.Text = "" Then
rs.Fields("Drives") = " "
Else
rs.Fields("Drives") = text5.Text & ""
End If
If text6.Text = "" Then
rs.Fields("Adapter Info") = " "
Else
  rs.Fields("Adapter Info") = text6.Text & ""
End If

    rs.Update

End Sub

Private Sub Command3_Click()
On Error Resume Next
If Check1.Value = 1 Then
List1.ListIndex = 0
List3.Clear
List4.Clear
Timer2.Enabled = True
Else
List1.ListIndex = 0
Timer1.Enabled = True
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Call reloadlist

DoEvents
Set db = OpenDatabase(App.Path & "\Computers.mdb")
Set rs = db.OpenRecordset("Computer_Info")
    ' Populate the list box
    Do Until rs.EOF
        List2.AddItem rs.Fields("Computer Name")
        List2.ItemData(List2.NewIndex) = rs.Fields("ID")
        
        rs.MoveNext
        
    Loop

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
frmmain.StatusBar1.Panels(1).Text = "Status: Please wait storing information on computer: " & List1.Text
Do Until List1.ListCount = 0
List2.Text = List1.Text
If List2.Text = List1.Text Then
List3.AddItem List1.Text & " - " & "Computer name is already in database"
List1.RemoveItem List1.ListIndex
Text7.Text = 0
Text8.Text = ""
ProgressBar1.Value = ProgressBar1.Value + 1
Else
Text7.Text = 0
Text8.Text = ""
Command1_Click
text1.Text = List1.Text
Call GetAllData
Command2_Click
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
ProgressBar1.Value = ProgressBar1.Value + 1
End If
Loop
DoEvents
frmmain.StatusBar1.Panels(1).Text = "Status:"
Call List_Save(List3, App.Path & "\Errors.ini")
DoEvents
If List3.ListCount = 0 Then
MsgBox "All Done, all computers have been entered into the database."
Else
frmmsg.Show
End If
Call frmmain.reloadlist
Unload Me
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
   frmmain.StatusBar1.Panels(1).Text = "Status: Checking..."
   Dim hostent_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   

If List1.Text = "" Then
   Else
   hostent_addr = gethostbyname(List1.Text)

If hostent_addr = 0 Then
List3.AddItem List1.Text & " - " & "Offline"
List1.RemoveItem List1.ListIndex
Timer2.Enabled = False
Timer3.Enabled = True
      frmmain.StatusBar1.Panels(1).Text = "Status:"
      Exit Sub
    Else
End If
   
   RtlMoveMemory host, hostent_addr, LenB(host)
   RtlMoveMemory hostip_addr, host.hAddrList, 4
   
   ReDim temp_ip_address(1 To host.hLength)
   RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
   
   For i = 1 To host.hLength
      ip_address = ip_address & temp_ip_address(i) & "."
   Next
   List4.AddItem List1.Text
Timer2.Enabled = False
Timer3.Enabled = True
End If
frmmain.StatusBar1.Panels(1).Text = "Status:"

End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If List4.ListCount = List1.ListCount Then
Timer3.Enabled = False
List1.ListIndex = 0
ProgressBar1.Max = List1.ListCount
Timer1.Enabled = True
Exit Sub
End If
List1.ListIndex = List1.ListIndex + 1
Timer3.Enabled = False
Timer2.Enabled = True

End Sub

Private Sub timercount_Timer()
On Error Resume Next
Label2.Caption = "Total Computers: " & List1.ListCount
End Sub
Private Sub sink_OnObjectReady(ByVal objWbemObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
Dim i As Integer
i = Text7.Text
Set adapter = GetObject("winmgmts:Win32_NetworkAdapterConfiguration=" & i & "")
Description = adapter.Description

text6.Text = text6.Text & Description & ENTER

If IsNull(adapter.MACAddress) Then
    text6.Text = text6.Text & "No MAC Address" & ENTER
    text6.Text = text6.Text & "" & ENTER
Else
    text6.Text = text6.Text & "Mac: " & adapter.MACAddress & ENTER
    text6.Text = text6.Text & "" & ENTER
End If
 

Text7.Text = i + 1
End Sub

