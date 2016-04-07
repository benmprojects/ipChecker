VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PingIPv4"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   11055
   End
   Begin VB.Timer Timer1 
      Left            =   6720
      Top             =   120
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Default         =   -1  'True
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   90
      Width           =   1215
   End
   Begin VB.TextBox txtNameOrIP 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Text            =   "localhost"
      Top             =   180
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Name or IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   210
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public godown As Integer
Private PingIPv4 As PingIPv4


Private Sub PingCheck(IP As String, Description As String)
    Set PingIPv4 = New PingIPv4
    
    Text1.Text = Text1.Text & vbCrLf & Now & " - Trying to ping " & IP

    With PingIPv4
            
            If .Ping(IP) Then
                Text1.Text = Text1.Text & " - Succuess, trip time " & CStr(.RoundTripTime) & "ms " & Description & " is up"
            Else
                Text1.Text = Text1.Text & " - Failed " & Description & " is down"
                godown = godown + 1
            End If
            'If .Reason <> PFR_BAD_IP Then Text1.Text = Text1.Text & CStr(.Status)
    End With
    
    Call savelog
    
End Sub

Private Sub savelog()
    Dim FileNum
    
    FileNum = FreeFile
    
    Open App.Path & "\log.txt" For Output As FileNum
    Print #FileNum, Text1.Text
    Close FileNum
    
End Sub


Private Sub cmdPing_Click()
 Call PingCheck("192.168.32.222", "Wetek")
    Call PingCheck("192.168.1.51", "Room")
End Sub

Private Sub Form_Load()
Dim dattimer

    dattimer = DateAdd("n", 2, Now)
    Timer1.Interval = 10000
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
Dim shell As WshShell
Dim lngReturnCode As Long
Dim strShellCommand As String
Dim dattimer

    If Now <= dattimer Then Exit Sub
    dattimer = DateAdd("n", 2, Now)
    
    
    Set shell = New WshShell
    strShellCommand = App.Path & "\ssh.bat"
    
    Call PingCheck("192.168.1.222", "Wetek")
    Call PingCheck("192.168.1.51", "Room")
    
    If godown = 2 Then
        Text1.Text = Text1.Text & vbCrLf & Now & " - No Clients are up server is shutting down!!!!"
        Call savelog
        lngReturnCode = shell.Run(strShellCommand, vbNormalFocus, vbTrue)
    End If
    
    godown = 0

End Sub

