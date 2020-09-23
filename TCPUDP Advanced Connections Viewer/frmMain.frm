VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "TCP/UDP Advanced Connections Viewer"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Applications"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Connections"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Current Connections"
         Height          =   4335
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   6735
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   5160
            TabIndex        =   10
            Top             =   3840
            Width           =   1455
         End
         Begin MSComctlLib.ListView lvNetstat 
            Height          =   3495
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   6165
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Protocol"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Local IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Local Port"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Remote IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Remote Port"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "State"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Running on Ports"
         Height          =   2055
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   6735
         Begin MSComctlLib.ListView lvApp2 
            Height          =   1695
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   2990
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Protocol"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Local IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Local Port"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Remote IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Remote Port"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "State"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Applications Accessing the internet"
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   6735
         Begin VB.Timer tCount 
            Interval        =   500
            Left            =   2880
            Top             =   960
         End
         Begin VB.PictureBox pic16 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   5340
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   4
            Top             =   1320
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox pic32 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            Height          =   480
            Left            =   4800
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   3
            Top             =   1320
            Visible         =   0   'False
            Width           =   480
         End
         Begin MSComctlLib.ListView lvApps 
            Height          =   1815
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   3201
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "ilLargeIcons"
            SmallIcons      =   "ilSmallIcons"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Application Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Version Number"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Application Path"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList ilSmallIcons 
      Left            =   3480
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilMenu 
      Left            =   4680
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0038
            Key             =   "set"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1098
            Key             =   "arc"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20EC
            Key             =   "sav"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5055
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilLargeIcons 
      Left            =   2640
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuPortOptions 
      Caption         =   "Port Options"
      Visible         =   0   'False
      Begin VB.Menu mnuClosePort 
         Caption         =   "Try Closeing Port"
      End
   End
   Begin VB.Menu MnuProgramOptions 
      Caption         =   "Program Options"
      Visible         =   0   'False
      Begin VB.Menu mnuTerminate 
         Caption         =   "Terminate Application"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'*************************************************************
'***                                                       ***
'***    Program Name: TCP/UDP Advanced Connection Viewer   ***
'***    Programmer:   Jake Paternoster (§e7eN)             ***
'***    Contact:      Hate_114@hotmail.com                 ***
'***    Date:         11:05 PM 5/10/2003                   ***
'***                                                       ***
'***    Description:  This program will show what progarms ***
'***                  are currently Connected/Listening    ***
'***                  on certain Ports.                    ***
'***                                                       ***
'***                  Please Comment and vote              ***
'*************************************************************
'*************************************************************

Private Sub cmdRefresh_Click()
    GetCurrentConnections
End Sub

Private Sub Form_Load()
    LvResize lvApps
    LvResize lvApp2
    LvResize lvNetstat
    GetCurrentConnections
    LoadRunningApps lvApps
End Sub

Private Sub lvApps_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim Connect() As Connection
Dim Item2 As ListItem
Dim lLength As Long

On Error Resume Next

lvApp2.ListItems.Clear

GetConnections Item.SubItems(2), lLength, Connect

For X = 0 To lLength - 1
    With Connect(X)
        If .Protocol = TCP Then
            Set Item2 = lvApp2.ListItems.Add(, , "TCP")
            Item2.SubItems(1) = .LocalIP
            Item2.SubItems(2) = .LocalPort
            Item2.SubItems(3) = .RemoteIP
            Item2.SubItems(4) = .RemotePort
            Item2.SubItems(5) = StateConvert(.state)
        Else
            Set Item2 = lvApp2.ListItems.Add(, , "UDP")
            Item2.SubItems(1) = .LocalIP
            Item2.SubItems(2) = .LocalPort
        End If

    End With
Next
End Sub

Private Sub lvApps_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 And lvApps.SelectedItem.Text <> "" Then
    Me.PopupMenu MnuProgramOptions
End If
End Sub

Private Sub lvNetstat_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 And lvNetstat.SelectedItem.Text <> "" Then
    Me.PopupMenu mnuPortOptions
End If
End Sub



Private Sub tCount_Timer()
    LoadRunningApps lvApps
    SB.SimpleText = "Last refreshed '" & Now & "'"
End Sub

Sub GetCurrentConnections()
Dim tUdpTable As MIB_UDPTABLE
Dim tTcpTable As MIB_TCPTABLE
Dim ldwSize As Long
Dim bOrder As Long

lvNetstat.ListItems.Clear
'==============================================================================
'                               GET UDP CONNECTIONS
'==============================================================================

    Call GetUdpTable(tUdpTable, ldwSize, bOrder) 'Call it once to get ldwSize
    Call GetUdpTable(tUdpTable, ldwSize, bOrder)

    'cycle for every connection in the table
    For X = 0 To tUdpTable.dwNumEntries - 1
        'Add it to the info into the listview box
        lvNetstat.ListItems.Add , , "UDP"
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(1) = IPconvert(tUdpTable.table(X).dwLocalAddr)
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(2) = PortConvert(tUdpTable.table(X).dwLocalPort)
    Next
    
'==============================================================================
'                               GET TCP CONNECTIONS
'==============================================================================
'Very similar to above, so just read them comments

    Call GetTcpTable(tTcpTable, ldwSize, bOrder)
    Call GetTcpTable(tTcpTable, ldwSize, bOrder)

    For X = 0 To tTcpTable.dwNumEntries - 1
        lvNetstat.ListItems.Add , , "TCP"
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(1) = IPconvert(tTcpTable.table(X).dwLocalAddr)
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(2) = PortConvert(tTcpTable.table(X).dwLocalPort)
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(3) = IPconvert(tTcpTable.table(X).dwRemoteAddr)
        
        If tTcpTable.table(X).dwState = 2 Then
            lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(4) = ""
        Else
            lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(4) = PortConvert(tTcpTable.table(X).dwRemotePort)
        End If
        
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(5) = StateConvert(tTcpTable.table(X).dwState)
    Next
End Sub

'==============================================================================
'                               Menu Items
'==============================================================================


Private Sub mnuClosePort_Click()
Dim Item As ListItem
    Set Item = lvNetstat.SelectedItem
    
    With Item
        ChangePortState UDP, .SubItems(2), DELETE_TCB
    End With
End Sub

Private Sub mnuTerminate_Click()
Dim Item As ListItem
Dim Ret As VbMsgBoxResult

    Set Item = lvApps.SelectedItem
    
    Ret = MsgBox("Do you wish to terminate '" & Item.Text & "'?", vbApplicationModal + vbYesNo + vbQuestion, "Terminate Application")
    If Ret = vbNo Then Exit Sub
    
    With Item
        For X = 0 To AppNo
            If Apps(X).Location = .SubItems(2) Then
                If TerminateProcessById(Apps(X).pID) Then MsgBox "Program Terminated Successfully!", vbApplicationModal + vbOKOnly + vbInformation, "Success!"
                Exit For
            End If
        Next
    End With
End Sub

Private Sub mnuAdd_Click()
Dim Item As ListItem
    Set Item = lvApps.SelectedItem
    
    With Item
        AddApp .SubItems(2)
    End With

End Sub

Private Sub lvApp2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 And lvApp2.SelectedItem.Text <> "" Then
    Me.PopupMenu mnuPortOptions
End If
End Sub
