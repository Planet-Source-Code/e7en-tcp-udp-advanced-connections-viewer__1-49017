VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      Picture         =   "FrmSplash.frx":0000
      ScaleHeight     =   4305
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.Timer tScroll 
         Interval        =   30
         Left            =   1200
         Top             =   1800
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   6705
         TabIndex        =   1
         Top             =   0
         Width           =   6735
         Begin VB.Label Nfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   45
         End
      End
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Str As String
Load frmMain

Str = "Long ago in a galaxy far far away..." & vbCrLf & vbCrLf & vbCrLf
Str = Str & "Hey guys, I felt like something should be written here." & vbCrLf
Str = Str & "Just looked like a waste of space. " & vbCrLf
Str = Str & "So what you think of the background eh?" & vbCrLf
Str = Str & "Amazing what a little screwing round in Photoshop can do. " & vbCrLf & vbCrLf
Str = Str & "Ok, spose I should tell ya something about this lil progam" & vbCrLf
Str = Str & "Basically gets the netstat details from the file," & vbCrLf
Str = Str & "reads them into variables then looks up all the process IDs" & vbCrLf
Str = Str & "and groups the connections into accordance with the programs" & vbCrLf
Str = Str & "running. also grabs the Icons and fileinfo. " & vbCrLf & vbCrLf
Str = Str & "Anyways enough shit from me, EnJoY and please Comment and Vote =)"

Nfo.Caption = Str

Nfo.Top = Picture2.Height
End Sub

Private Sub Picture1_Click()
frmMain.Show
Unload Me
End Sub

Private Sub tScroll_Timer()
    Nfo.Top = Nfo.Top - 5
End Sub
