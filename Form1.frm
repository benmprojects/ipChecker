VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Checker"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   12000
      Left            =   720
      Top             =   4200
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   10455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetRTTAndHopCount _
    Lib "iphlpapi.dll" _
   (ByVal lDestIPAddr As Long, _
    ByRef lHopCount As Long, _
    ByVal lMaxHops As Long, _
    ByRef lRTT As Long) As Long
        
Private Declare Function inet_addr _
    Lib "wsock32.dll" _
   (ByVal cp As String) As Long
Public godown As Integer
Private datTimer As Date

Public Function Ping(prmIPaddr As String) As Boolean
Dim IPaddr As Long, HopsCount As Long, RTT As Long
Dim MaxHops As Long
    Const SUCCESS = 1
    MaxHops = 20
    IPaddr = inet_addr(prmIPaddr)
    Ping = (GetRTTAndHopCount(IPaddr, HopsCount, MaxHops, RTT) = SUCCESS)
End Function

Private Sub Pingcheck(IP As String, Description As String)
Dim result

    Text1.Text = Text1.Text & vbCrLf & Now & " - Trying to ping " & IP
    
    result = Ping(IP)
    
    If result = True Then
        Text1.Text = Text1.Text & " - Succuess " & Description & " is up"
    Else
        Text1.Text = Text1.Text & " - Failed " & Description & " is down"
        godown = godown + 1
    End If
    
    Call savelog

End Sub



Private Sub savelog()
    FileNum = FreeFile
    
    Open App.Path & "\log.txt" For Output As FileNum
    Print #FileNum, Text1.Text
    Close FileNum
    
End Sub


Private Sub Form_Load()
    datTimer = DateAdd("n", 2, Now)
    Timer1.Interval = 10000
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim shell As WshShell
Dim lngReturnCode As Long
Dim strShellCommand As String

    If Now <= datTimer Then Exit Sub
    datTimer = DateAdd("n", 2, Now)
    
    
    Set shell = New WshShell
    strShellCommand = App.Path & "\ssh.bat"
    
    Call Pingcheck("192.168.1.49", "Wetek")
    Call Pingcheck("192.168.1.51", "Room")
    
    If godown = 2 Then
        Text1.Text = Text1.Text & vbCrLf & Now & " - No Clients are up server is shutting down!!!!"
        Call savelog
        lngReturnCode = shell.Run(strShellCommand, vbNormalFocus, vbTrue)
    End If
    
    godown = 0

End Sub

