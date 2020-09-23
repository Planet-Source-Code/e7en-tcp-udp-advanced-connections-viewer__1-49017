Attribute VB_Name = "modNetstat2"
Option Explicit

Declare Function GetTcpTable Lib "IPhlpAPI" (pTcpTable As MIB_TCPTABLE, pdwSize As Long, bOrder As Long) As Long
Declare Function GetUdpTable Lib "IPhlpAPI" (pUdpTable As MIB_UDPTABLE, pdwSize As Long, bOrder As Long) As Long
Declare Function SetTcpEntry Lib "IPhlpAPI" (pTcpRow As MIB_TCPROW) As Long 'This is used to close an open port.

Type MIB_TCPROW
    dwState As Long              'state of the connection
    dwLocalAddr As String * 4    'address on local computer
    dwLocalPort As String * 4    'port number on local computer
    dwRemoteAddr As String * 4   'address on remote computer
    dwRemotePort As String * 4   'port number on remote computer
End Type

Type MIB_TCPTABLE
    dwNumEntries As Long         'number of entries in the table
    table(100) As MIB_TCPROW     'array of TCP connections
End Type

Type MIB_UDPROW
    dwLocalAddr As String * 4    'address on local computer
    dwLocalPort As String * 4    'port number on local computer
End Type

Type MIB_UDPTABLE
    dwNumEntries As Long         'number of entries in the table
    table(100) As MIB_UDPROW     'table of MIB_UDPROW structs
End Type

Sub ChangePortState(Protocol As Protocol, LocPort As Long, state As state)
Dim tUdpTable As MIB_UDPTABLE
Dim tTcpTable As MIB_TCPTABLE
Dim ldwSize As Long
Dim bOrder As Long
Dim X As Long

'==============================================================================
'                               GET TCP CONNECTIONS
'==============================================================================
'Very similar to above, so just read them comments

    Call GetTcpTable(tTcpTable, ldwSize, bOrder)
    Call GetTcpTable(tTcpTable, ldwSize, bOrder)

    For X = 0 To tTcpTable.dwNumEntries - 1
            If PortConvert(tTcpTable.table(X).dwLocalPort) = LocPort Then
                tTcpTable.table(X).dwState = state
                SetTcpEntry tTcpTable.table(X)
            Exit For
        End If
    Next
    
End Sub

Function IPconvert(sIP) As String
Dim X As Integer
    
    'convert the string into an IP
    For X = 1 To Len(sIP)
        IPconvert = IPconvert & Asc(Mid(sIP, X, 1)) & "."
    Next
    
'remove last '.'
IPconvert = Left(IPconvert, Len(IPconvert) - 1)
End Function

Function PortConvert(sPort) As String
Dim lPort As Long

'Convert string into the port number
lPort = Asc(Mid(sPort, 1, 1))
lPort = lPort * 256
lPort = lPort + Asc(Mid(sPort, 2, 1))

PortConvert = lPort
End Function

Function StateConvert(sState) As String

'Convert the number into the corresponding port status
Select Case sState - 1
    Case 0
        StateConvert = "CLOSED"
    Case 1
        StateConvert = "LISTENING"
    Case 2
        StateConvert = "SYN_SENT"
    Case 3
        StateConvert = "SYN_RCVD"
    Case 4
        StateConvert = "ESTABLISHED"
    Case 5
        StateConvert = "FIN_WAIT1"
    Case 6
        StateConvert = "FIN_WAIT2"
    Case 7
        StateConvert = "CLOSE_WAIT"
    Case 8
        StateConvert = "CLOSING"
    Case 9
        StateConvert = "LAST_ACK"
    Case 10
        StateConvert = "TIME_WAIT"
    Case 11
        StateConvert = "DELETE_TCB"
    Case Else
        StateConvert = "UNKNOWN"
End Select
End Function


