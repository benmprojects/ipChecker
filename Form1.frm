VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IsUp"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11250
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
   ScaleHeight     =   3090
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7440
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
   Begin VB.Timer Timer1 
      Left            =   6720
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public godown As Integer
Public dattimer
Public notif
Public dattimer2
Public sSetting1 As String
Public sSetting2 As String
Private PingIPv4 As PingIPv4
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public runtimeex

Private Sub PingCheck(IP As String, Description As String)
    Set PingIPv4 = New PingIPv4
    
    Text1.SelText = vbCrLf & Now & " - Trying to ping " & IP

    With PingIPv4
            
            If .Ping(IP) Then
                Text1.SelText = " - Succuess, trip time " & CStr(.RoundTripTime) & "ms " & Description & " is up"
            Else
                Text1.SelText = " - Failed " & Description & " is down"
                godown = godown + 1
            End If
            'If .Reason <> PFR_BAD_IP Then Text1.Text = Text1.Text & CStr(.Status)
    End With
    
    Call savelog
    
End Sub

Private Sub savelog()
    Dim FileNum
    
    FileNum = FreeFile
    
    Open App.Path & "\log.txt" For Append As FileNum
    Print #FileNum, Text1.Text
    Close FileNum
    
End Sub

Private Function Notification(ByVal URL As String) As String
    Dim Ans As String
    Dim oHTTP As MSXML2.XMLHTTP, sBuffer As String
    On Error Resume Next
    Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    oHTTP.Open "GET", URL, False
    oHTTP.send
    sBuffer = oHTTP.responseText
    Set oHTTP = Nothing
    
    If sBuffer = "" Then
        Notification "ERROR no response"
    Else
        Notification = sBuffer
    End If

End Function

Private Sub Form_Load()
    
    sSetting1 = GetINISetting("Wetek", "IP", App.Path & "\SETTINGS.INI")
    sSetting2 = GetINISetting("Room", "IP", App.Path & "\SETTINGS.INI")
    
    
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = Now & " - IsUp is Running"
    Text1.SelText = vbCrLf & Now & " - Sending notification to Kodi clients "
    notif = Notification("http://" & sSetting1 & "/jsonrpc?request={""jsonrpc"":""2.0"",""method"":""GUI.ShowNotification"",""params"":{""title"":""IsUp"",""message"":""Server IsUp""},""id"":1}")
    Text1.SelText = vbCrLf & Now & " - JSON response from  " & sSetting1 & " " & notif
    notif = Notification("http://" & sSetting2 & "/jsonrpc?request={""jsonrpc"":""2.0"",""method"":""GUI.ShowNotification"",""params"":{""title"":""IsUp"",""message"":""Server IsUp""},""id"":1}")
    Text1.SelText = vbCrLf & Now & " - JSON response from " & sSetting2 & " " & notif
    Call savelog
    dattimer = DateAdd("n", 2, Now)
    runtimeex = DateAdd("n", 180, Now)
    Timer1.Interval = 10000
    Timer1.Enabled = True

    
End Sub

Private Sub Timer1_Timer()
Dim oshell As WshShell
Dim ShellCommand As Long
Dim strShellCommand1 As String
'Dim lMinutes As Long


 
    Set oshell = New WshShell
    
    'Text1.Text = Text1.Text & vbCrLf & Now & " The server has been up for " & AppRunTime - Now
    

    If DateDiff("n", Now, runtimeex) < 0 Then
        godown = 2
        Timer2.Enabled = True
    End If

    If Now <= dattimer Then Exit Sub
    dattimer = DateAdd("n", 2, Now)
    
  
    Call PingCheck(sSetting1, "Wetek")
    Call PingCheck(sSetting2, "Room")
    
    If godown = 2 Then
        Text1.SelText = vbCrLf & Now & " - No Clients are up, server is shutting down!!!!"
        Text1.SelText = vbCrLf & Now & " Sending notification to Kodi clients "
        notif = Notification("http://" & sSetting1 & "/jsonrpc?request={""jsonrpc"":""2.0"",""method"":""GUI.ShowNotification"",""params"":{""title"":""IsUp"",""message"":""No Clients are up, server is shutting down!!!!""},""id"":1}")
        Text1.SelText = vbCrLf & Now & " - JSON response from " & sSetting1 & " " & notif
        notif = Notification("http://" & sSetting2 & "/jsonrpc?request={""jsonrpc"":""2.0"",""method"":""GUI.ShowNotification"",""params"":{""title"":""IsUp"",""message"":""No Clients are up, server is shutting down!!!!""},""id"":1}")
        Text1.SelText = vbCrLf & Now & " - JSON response from " & sSetting2 & " " & notif
        Call savelog
       ' ShellCommand = oshell.Run("C:\WINDOWS\system32\shutdown.exe -s -t 0", vbNormalFocus, vbTrue)
        dattimer2 = DateAdd("n", 1, Now)
        Timer2.Enabled = True
        Timer1.Enabled = False
    End If
    
    godown = 0

End Sub


Private Sub Timer2_Timer()
Dim shell As WshShell
Dim lngReturnCode As Long
Dim strShellCommand As String

    Set shell = New WshShell

    If Now <= dattimer2 Then Exit Sub
    
    Text1.SelText = vbCrLf & Now & " Windows Shutdown failed. Hard shut down by DRAC Controller"
    Text1.SelText = vbCrLf & Now & " Sending notification to Kodi clients "
    notif = Notification("http://" & sSetting1 & "/jsonrpc?request={""jsonrpc"":""2.0"",""method"":""GUI.ShowNotification"",""params"":{""title"":""IsUp"",""message"":""Windows Shutdown failed. Hard shut down by DRAC Controller""},""id"":1}")
    Text1.SelText = vbCrLf & Now & " JSON response from " & sSetting1 & " " & notif
    notif = Notification("http://" & sSetting2 & "/jsonrpc?request={""jsonrpc"":""2.0"",""method"":""GUI.ShowNotification"",""params"":{""title"":""IsUp"",""message"":""Windows Shutdown failed. Hard shut down by DRAC Controller""},""id"":1}")
    Text1.SelText = vbCrLf & Now & " JSON response from " & sSetting2 & " " & notif
    
    lngReturnCode = shell.Run(App.Path & "\putty.exe -ssh root@192.168.1.54 22 -pw Ret5aM321 -m  " & App.Path & "\command.txt", vbNormalFocus, vbTrue)
    
End Sub
