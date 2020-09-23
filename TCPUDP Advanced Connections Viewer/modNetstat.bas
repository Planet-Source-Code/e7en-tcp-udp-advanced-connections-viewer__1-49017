Attribute VB_Name = "modNetstat"
Enum Protocol
    TCP = 0
    UDP = 1
End Enum

Enum state
    CLOSED = 1
    LISTENING = 2
    SYN_SENT = 3
    SYN_RCVD = 4
    ESTABLISHED = 5
    FIN_WAIT1 = 6
    FIN_WAIT2 = 7
    CLOSE_WAIT = 8
    CLOSING = 9
    LAST_ACK = 10
    TIME_WAIT = 11
    DELETE_TCB = 12
End Enum


Type Connection
    Protocol As Protocol
    LocalIP As String
    LocalPort As Long
    RemoteIP As String
    RemotePort As Long
    state As state
    pID As Long
    ProcessName As String
End Type

Type App
    Name As String
    Location As String
    Version As String
    Suspended As Boolean
    pID As Long
    ConnectionCount As Long
    Connections(50) As Connection
End Type

Public Apps() As App
Public AppNo As Long
Dim Connections() As Connection
Dim FileOpen As Boolean

Function FileExist(sfile As String) As Boolean
    If Dir(sfile) <> "" Then FileExist = True Else FileExist = False
End Function

Function GetConnections(sAppPath As String, lLength As Long, Connect() As Connection)

For X = 1 To AppNo
    If Apps(X).Location = sAppPath Then
        lLength = Apps(X).ConnectionCount
        Connect = Apps(X).Connections
        Exit For
    End If
Next
End Function

Function GetNetstat() As String
Dim f As Long
Dim sBuff As String
Dim sfile As String

 sfile = "c:\Tmp$.dat"
 
 If FileOpen = True Then Exit Function
 'If FileExist(sfile) = True Then Kill sfile

 Call Shell("command.com /c netstat -an -o > " & sfile, vbNormal)
 
 'Do Until FileExist(sfile) = True
'    DoEvents
 'Loop
 
 f = FreeFile
 
 Open sfile For Input As #f
 FileOpen = True
    Do Until sBuff <> ""
        DoEvents
        sBuff = Input(LOF(f), f)
    Loop
 Close #f
 FileOpen = False
 
GetNetstat = sBuff
End Function

Sub LoadRunningApps(lv As ListView)
Dim Item As ListItem
Dim sFiles2() As String
Dim lCount As Long
Dim IconIndex As Integer
Dim LI As ListImage
Static lOld As Long
Dim r As Integer

ReadConnections

If AppNo <> lOld Then
    lOld = AppNo
    lv.ListItems.Clear
    'lv.Icons = frmMain.ilLargeIcons
    'lv.SmallIcons = frmMain.ilSmallIcons
    
    For X = 1 To AppNo
        Set Item = lv.ListItems.Add(, , Apps(X).Name)
        Item.SubItems(1) = Apps(X).Version
        Item.SubItems(2) = Apps(X).Location
    Next X
    
    For X = 1 To lv.ListItems.Count - 1
        
    r = ExtractIcon(lv.ListItems(X).SubItems(2), frmMain.ilSmallIcons, frmMain.pic16, 16)
    
    If r = 0 Then
        MsgBox "Error Loading Icon!", vbCritical + vbApplicationModal + vbOKOnly, "Error"
    Else
        lv.ListItems(X).SmallIcon = r
    End If
    
    r = ExtractIcon(lv.ListItems(X).SubItems(2), frmMain.ilLargeIcons, frmMain.pic32, 32)
    
    If r = 0 Then
        MsgBox "Error Loading Icon!", vbCritical + vbApplicationModal + vbOKOnly, "Error"
    Else
        lv.ListItems(X).Icon = r
    End If
    
Next
End If

lv.Refresh
End Sub

Sub ReadConnections()
Dim sData As String
Dim iState As Integer
Dim sSplitData() As String
Dim Item As ListItem
Dim X As Integer
Dim lCount As Long
Dim Res As FILEPROPERTIE
Dim Unique As Boolean

sData = GetNetstat

ReDim Apps(100)
 'Open "c:\Tmp$.dat" For Input As #1
 '       sData = Input(LOF(1), 1)
 'Close #1

Do While InStr(1, sData, "  ")
    sData = Replace(sData, "  ", " ")
    DoEvents
Loop

If sData = "" Then Exit Sub
sSplitData = Split(sData, vbCrLf)

ReDim Connections(UBound(sSplitData))
GetProcesses

For X = 4 To UBound(sSplitData) - 1
    With Connections(X)
    
    If Split(sSplitData(X), " ")(1) = "TCP" Or Split(sSplitData(X), " ")(1) = "UDP" Then
        
        If Split(sSplitData(X), " ")(1) = "TCP" Then
            .Protocol = 0
        Else
            .Protocol = 1
        End If
        
        .LocalIP = Split(Split(sSplitData(X), " ")(2), ":")(0)
        .LocalPort = Split(Split(sSplitData(X), " ")(2), ":")(1)
        
        If .Protocol = 0 Then
            .RemoteIP = Split(Split(sSplitData(X), " ")(3), ":")(0)
            .RemotePort = Split(Split(sSplitData(X), " ")(3), ":")(1)
            iState = StateConv(Split(sSplitData(X), " ")(4))
            .state = iState
            .pID = Split(sSplitData(X), " ")(5)
        Else
            .pID = Split(sSplitData(X), " ")(4)
        End If
        .ProcessName = GetProcessName(.pID)
        
        If .ProcessName <> "" Then
            Res = FileInfo(.ProcessName)
            
            For y = 1 To lCount
                If Apps(y).Location = .ProcessName Then
                    For Z = 1 To Apps(y).ConnectionCount
                        If Apps(y).Connections(Z).LocalPort = .LocalPort Then
                            Unique = False
                            Exit For
                        Else
                            Unique = True
                        End If
                    Next Z
                    
                    If Unique = True Then
                        Apps(y).Connections(Apps(y).ConnectionCount) = Connections(X)
                        Apps(y).ConnectionCount = Apps(y).ConnectionCount + 1
                    End If
                            
                    Unique = False
                    Exit For
                Else
                    Unique = True
                End If
            Next y
            
            If Unique = True Or lCount = 0 Then
                Apps(lCount).Name = Res.ProductName
                Apps(lCount).Location = .ProcessName
                Apps(lCount).pID = .pID
                Apps(lCount).Version = Res.ProductVersion
                lCount = lCount + 1
                Apps(y).Connections(Apps(y).ConnectionCount) = Connections(X)
                Apps(y).ConnectionCount = Apps(y).ConnectionCount + 1
            End If
        End If
    End If
    End With
Next X

AppNo = lCount

'For x = 4 To UBound(Connections) - 1
'
'    With Connections(x)
'        Set Item = lvConnections.ListItems.Add(, , .Protocol)
'        Item.SubItems(1) = .LocalIP
'        Item.SubItems(2) = .LocalPort
'        Item.SubItems(3) = .RemoteIP
'        Item.SubItems(4) = .RemotePort
'        Item.SubItems(5) = .state
'        Item.SubItems(6) = .pID
'        Item.SubItems(7) = .ProcessName
'    End With
'Next x
End Sub

Function GetProcessName(pID As Long) As String
    For X = 1 To UBound(Procs)
        If Procs(X).pID = pID Then
            GetProcessName = Procs(X).ProcessName
            Exit Function
        End If
    Next X
End Function

Private Function StateConv(sState) As Integer

Select Case Trim(UCase(sState))
  Case "UNKNOWN": StateConv = 0
  Case "CLOSED": StateConv = 1
  Case "LISTENING": StateConv = 2
  Case "SYN_SENT": StateConv = 3
  Case "SYN_RCVD": StateConv = 4
  Case "ESTABLISHED": StateConv = 5
  Case "FIN_WAIT1": StateConv = 6
  Case "FIN_WAIT2": StateConv = 7
  Case "CLOSE_WAIT": StateConv = 8
  Case "CLOSING": StateConv = 9
  Case "LAST_ACK": StateConv = 10
  Case "TIME_WAIT": StateConv = 11
  Case "DELETE_TCB":  StateConv = 12
End Select

End Function

