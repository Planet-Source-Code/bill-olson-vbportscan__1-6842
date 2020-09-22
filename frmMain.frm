VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Port Scanner"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBanners 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   5520
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1800
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "5000"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtToPort 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "65355"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtFromPort 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "1"
      Top             =   1320
      Width           =   615
   End
   Begin VB.ListBox lstClosedPorts 
      Height          =   1035
      Left            =   3000
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.ListBox lstOpenPorts 
      Height          =   1035
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Port Banners:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Open Ports:"
      Height          =   195
      Left            =   3000
      TabIndex        =   15
      Top             =   240
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Closed Ports:"
      Height          =   195
      Left            =   3000
      TabIndex        =   14
      Top             =   1680
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Time Out (msec):"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To:"
      Height          =   195
      Left            =   1440
      TabIndex        =   12
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "From:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ports to Scan:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remote Host (URL or IP Address):"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   2430
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private gblnStop As Boolean
Private gblnExit As Boolean
Private gblnResume As Boolean
Private gblnBannerGrabbed As Boolean

Private Type Banner
    Port As Long
    BannerText As String
End Type

Private gtypBanners() As Banner

Dim glngPort As Long

Private Sub cmdExit_Click()
    gblnStop = True
    gblnExit = True
    Unload Me
End Sub

Private Sub cmdScan_Click()
On Error GoTo ErrorHandler
    Dim l As Long
    Dim x As Long
    Dim i As Integer
    Dim strBanner As String
    Dim strData As String
        
    ReDim gtypBanners(0)
    txtBanners.Text = ""
        
    If Not IsNumeric(txtTimeOut.Text) Then
        MsgBox "Timeout must be numeric.", vbExclamation
        txtTimeOut.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtFromPort.Text) Then
        MsgBox "Port number must be numeric.", vbExclamation
        txtFromPort.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtToPort.Text) Then
        MsgBox "Port number must be numeric.", vbExclamation
        txtToPort.SetFocus
        Exit Sub
    End If
    
    gblnStop = False
    
    lstOpenPorts.Clear
    lstClosedPorts.Clear
    lblStatus.Caption = ""
    
    If Winsock1.State <> 0 Then
        Winsock1.Close
    End If
    
    cmdStop.Enabled = True
    cmdScan.Enabled = False
    
    For glngPort = txtFromPort.Text To txtToPort.Text
        l = GetTickCount()
        lblStatus.Caption = "Scanning port " & glngPort
        Winsock1.Connect txtRemoteHost.Text, glngPort
        Do While True
            gblnResume = False
            Select Case Winsock1.State
                Case sckClosed
                    lblStatus.Caption = "Port " & glngPort & " closed"
                Case sckOpen
                    lblStatus.Caption = "Port " & glngPort & " open"
                Case sckListening
                    lblStatus.Caption = "Listening"
                Case sckConnectionPending
                    lblStatus.Caption = "Connection pending on port " & glngPort
                Case sckResolvingHost
                    lblStatus.Caption = "Resolving Host"
                Case sckHostResolved
                    lblStatus.Caption = "Host Resolved"
                Case sckConnecting
                    lblStatus.Caption = "Connecting to port " & glngPort
                Case sckConnected
                    gblnResume = True
                    lblStatus.Caption = "Connected to port " & glngPort
                Case sckClosing
                    lblStatus.Caption = "Closing port " & glngPort
                Case sckError
                    'lblStatus.Caption = "Error connecting to port " & glngPort & " (potentially closed)"
            End Select
            x = GetTickCount()
            If x - l >= txtTimeOut.Text Then
                DoEvents
                Exit Do
            End If
            
            If gblnStop = True Then
                Exit Do
            End If
            DoEvents
            If gblnResume = True Then Exit Do
        Loop
        
        If Winsock1.State <> 0 Then
            Winsock1.Close
        End If
        
        If gblnStop = True Then
            Exit For
        End If
    Next
    lblStatus.Caption = "Ready"
    cmdStop.Enabled = False
    If gblnExit = True Then
        End
    End If
    cmdScan.Enabled = True
Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.number & " - " & Err.Description, vbCritical, "Error Occurred", Err.HelpFile, Err.HelpContext
End Sub

Private Sub cmdStop_Click()
    gblnStop = True
    cmdStop.Enabled = False
    lblStatus.Caption = "Ready"
    DoEvents
End Sub

Private Sub Form_Load()
    txtRemoteHost.Text = Winsock1.LocalIP
    DoEvents
End Sub

Private Sub Winsock1_Connect()
    On Error Resume Next
    lstOpenPorts.AddItem glngPort
    If Winsock1.BytesReceived = 0 Then
        If Winsock1.State <> 0 Then
            Winsock1.SendData vbCrLf
            Idle 500
            If Winsock1.BytesReceived = 0 Then
                If Winsock1.State <> 0 Then
                    Winsock1.SendData vbCrLf
                    Idle 500
                End If
            End If
        End If
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim blnExists As Boolean
    Dim i As Long
    
    Winsock1.GetData strData, , bytesTotal
    
    For i = 1 To UBound(gtypBanners)
        If gtypBanners(i).Port = glngPort Then
            blnExists = True
        End If
    Next
    
    If blnExists = False Then
        ReDim Preserve gtypBanners(UBound(gtypBanners) + 1)
        gtypBanners(UBound(gtypBanners)).Port = glngPort
        gtypBanners(UBound(gtypBanners)).BannerText = strData
        txtBanners.Text = txtBanners.Text & "Port: "
        txtBanners.Text = txtBanners.Text & gtypBanners(UBound(gtypBanners)).Port & vbCrLf
        txtBanners.Text = txtBanners.Text & gtypBanners(UBound(gtypBanners)).BannerText & vbCrLf
        txtBanners.Text = txtBanners.Text & "---------------------------------" & vbCrLf
    End If
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Select Case number
        Case 10061
            lstClosedPorts.AddItem glngPort
            lblStatus.Caption = "Port " & glngPort & " Closed "
            gblnResume = True
            DoEvents
        Case Else
            MsgBox "Error: " & number & " - " & Description & " - " & Source & " - " & Scode, vbCritical, "Error Occurred"
    End Select
    DoEvents
End Sub

Public Sub Idle(ByVal lngMsec As Long)
    Dim lngCurrent As Long
    Dim lngPast As Long

    lngPast = GetTickCount
    
    Do Until lngCurrent - lngPast >= lngMsec
        If gblnExit = True Then Exit Sub
        lngCurrent = GetTickCount
        DoEvents
    Loop
End Sub
