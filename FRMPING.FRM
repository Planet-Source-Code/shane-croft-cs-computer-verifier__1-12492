VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmping 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check computer Offline/Online Status"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   Icon            =   "frmping.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1320
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   5520
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   5520
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   6345
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check All"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8493
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Computer Name"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Computers: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select a Computer Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private db As Database
Private rs As Recordset
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

Private Sub Command1_Click()
   On Error Resume Next
   frmmain.StatusBar1.Panels(1).Text = "Status: Checking..."
   Dim hostent_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   
If Combo1.Text = "" Then
   Else
   hostent_addr = gethostbyname(Combo1.Text)

If hostent_addr = 0 Then
    Set Item = ListView1.ListItems.Add(, , Combo1.Text)
    Item.SubItems(1) = "OffLine"
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
   ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
    Set Item = ListView1.ListItems.Add(, , Combo1.Text)
    Item.SubItems(1) = "Online"
End If
frmmain.StatusBar1.Panels(1).Text = "Status:"

End Sub

Private Sub Command2_Click()
On Error Resume Next
ListView1.ListItems.Clear
List1.Clear
Combo1.ListIndex = 0
ProgressBar1.Value = 0
ProgressBar1.Max = Combo1.ListCount
Timer3.Enabled = True

End Sub

Private Sub Form_Load()
On Error Resume Next
    Set db = OpenDatabase(App.Path & "\Computers.mdb")
    Set rs = db.OpenRecordset("SELECT * FROM Computer_Info " & "ORDER BY [Computer Name]")
    
    ' Populate the list box
    Do Until rs.EOF
        Combo1.AddItem rs.Fields("Computer Name")
        Combo1.ItemData(List1.NewIndex) = rs.Fields("ID")
        
        rs.MoveNext
        
    Loop

End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Total Computers: " & Combo1.ListCount
If Combo1.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

If Combo1.ListCount = 0 Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If List1.ListCount = Combo1.ListCount Then
Timer2.Enabled = False
Exit Sub
End If
Combo1.ListIndex = Combo1.ListIndex + 1
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
   frmmain.StatusBar1.Panels(1).Text = "Status: Checking..."
   Dim hostent_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   
If Combo1.Text = "" Then
   Else
   hostent_addr = gethostbyname(Combo1.Text)

If hostent_addr = 0 Then
    Set Item = ListView1.ListItems.Add(, , Combo1.Text)
    Item.SubItems(1) = "OffLine"
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
   ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
    Set Item = ListView1.ListItems.Add(, , Combo1.Text)
    Item.SubItems(1) = "Online"
End If
List1.AddItem Combo1.Text
frmmain.StatusBar1.Panels(1).Text = "Status:"
ProgressBar1.Value = ProgressBar1.Value + 1
Timer3.Enabled = False
Timer2.Enabled = True
End Sub
