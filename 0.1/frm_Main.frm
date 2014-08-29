VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ForceHole For Blackhole v1.2.3 - Xyl2k!"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "URLs found"
      Height          =   2775
      Left            =   6840
      TabIndex        =   32
      Top             =   120
      Width           =   3615
      Begin VB.TextBox Text2 
         Height          =   2175
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load hole.txt"
      Height          =   255
      Left            =   5520
      TabIndex        =   29
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox List_Url 
      Height          =   2010
      ItemData        =   "frm_Main.frx":0442
      Left            =   5520
      List            =   "frm_Main.frx":0A88
      TabIndex        =   28
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Download 
      Caption         =   "Download"
      Default         =   -1  'True
      Height          =   255
      Left            =   8640
      TabIndex        =   24
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5160
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.CheckBox af 
      Caption         =   "After download open destination folder"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   6480
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   6840
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Download File"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         Picture         =   "frm_Main.frx":17DE
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Download"
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txt_Url 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Url:"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3255
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17637
            MinWidth        =   17637
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transfer Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1000
      Width           =   5295
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   4800
         Top             =   240
      End
      Begin VB.CommandButton cmd_Abort 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label UrlToDownload 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Processing:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbl_File 
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lbl_Loaded 
         Caption         =   "0 B / 0 B"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label lbl_Remain 
         Caption         =   "00:00:00"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lbl_Speed 
         Caption         =   "0 B/s"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lbl_CaptionRemain 
         Caption         =   "Remaining:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lbl_Done 
         Alignment       =   2  'Center
         Caption         =   "0%"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl_CaptionLoaded 
         Caption         =   "Downloaded:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lbl_CaptionRate 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lbl_CaptionFile 
         Caption         =   "File:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Label SampleNumber 
      Alignment       =   2  'Center
      Caption         =   "[0/534]"
      Height          =   255
      Left            =   5520
      TabIndex        =   30
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label FnameAdr 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label random 
      Caption         =   "Label3"
      Height          =   375
      Left            =   8280
      TabIndex        =   22
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Label3"
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing :"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Adr 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   5760
      Width           =   3135
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tDownload
    strFile As String
    strHost As String
    strPort As String
    strRequest As String
    lngLength As Long
    lngRec As Long
    lngStart As Long
    lngStatusTime As Long
    F As Integer
End Type

Dim Download As tDownload
Dim a, B As Integer
Dim ligne As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit

Private Sub Form_Load()
 centerform Me
 Randomize
 random.Caption = Int(1000 * Rnd)
 StatusBar1.Panels(1).Text = "» That should do the trick !"
End Sub

 Public Sub centerform(frm As Form) 'center form
 frm.Top = Screen.Height / 2 - frm.Height / 2
 frm.Left = Screen.Width / 2 - frm.Width / 2
 End Sub
 
Private Sub cmd_Abort_Click()
    On Error Resume Next
    Timer1.Enabled = False
    Socket.Close
    StatusBar1.Panels(1).Text = "Status: Canceled"
    Close #Download.F
    Kill Download.strFile
    txt_Url.Enabled = True
Load.Enabled = True
cmd_Abort.Enabled = False
End Sub

Private Sub cmd_Download_Click()
    Dim strFile As String
    Dim I As Integer
Dim strnew As String

Text1.Text = Time & "-" & random.Caption & ".Temp.ViR" 'random who is not random
strnew = Replace(Text1.Text, ":", "-")
Label1.Caption = strnew

Adr.Caption = UrlToDownload.Caption
FnameAdr.Caption = UrlToDownload.Caption
Adr.Caption = Replace(Adr.Caption, "http://", "")
Adr.Caption = Replace(Adr.Caption, "www.", "")
Adr.Caption = Replace(Adr.Caption, ":", "-")
Adr.Caption = Decoupe(Adr.Caption, "/", True, 1)
Adr.Caption = Decoupe(Adr.Caption, "/", True, 1)
Adr.Caption = Decoupe(Adr.Caption, "/", True, 1)
If Not Rep2 Then MkDir (App.path & "\" & Adr.Caption)

        Download.strFile = App.path & "\" & Adr & "\" & Label1.Caption
        Download.F = FreeFile
        Open Download.strFile For Binary As #Download.F
        
        'Start Download

StatusBar1.Panels(1).Text = "Downloading: " & UrlToDownload.Caption

FnameAdr.Caption = RcpNom(FnameAdr)
FnameAdr.Caption = Replace(FnameAdr.Caption, "=", "-")
FnameAdr.Caption = Replace(FnameAdr.Caption, "?", "-")

    On Error Resume Next
        ParseUrl UrlToDownload, Download.strHost, Download.strPort, Download.strRequest
        If Not Val(Download.strPort) = 0 Then
            'Url is valid download
            Download.lngRec = 0
            Socket.Close
            Socket.Connect Download.strHost, Download.strPort
            lbl_File.Caption = FnameAdr
        Else
            MsgBox "Error: The entered URL is not valid", vbOKOnly + vbExclamation, "Error"
        End If
End Sub

Private Sub Command1_Click()
Dim shll32 As New Shell
If txt_Url.Text = "" Then
StatusBar1.Panels(1).Text = "Give me your blackhole url !"
Exit Sub
Else
End If

If a = List_Url.ListCount Then 'we have downloaded all the shit?
a = 0
StatusBar1.Panels(1).Text = "Download(s) done !"
Timer1.Enabled = False
txt_Url.Enabled = True
Load.Enabled = True
cmd_Abort.Enabled = False
If af.Value = 1 Then
shll32.Explore App.path & "\" & Adr
Else
End If
Exit Sub
End If
B = a + 1
If a < List_Url.ListCount Then
UrlToDownload.Caption = txt_Url.Text & List_Url.List(a)
List_Url.Selected(a) = True
a = a + 1
Else
a = 0
Exit Sub
End If
SampleNumber.Caption = "[" & B & "/" & List_Url.ListCount & "]"
txt_Url.Enabled = False
Load.Enabled = False
cmd_Abort.Enabled = True
'Start Download
StatusBar1.Panels(1).Text = "Downloading: " & UrlToDownload.Caption
Call cmd_Download_Click
End Sub

 Public Function RcpNom(Adresse As String, Optional AcExt As Boolean = True) As String
 If (Len(Adresse) < 4) Or (InStr(Adresse, "/") = 0) Then Exit Function
 If AcExt Then RcpNom = Decoupe(Adresse, "/") Else RcpNom = Decoupe(Decoupe(Adresse, "/"), ".", True)
 End Function 'url decoupe
 
 Private Function Decoupe(Chn As String, Chr As String, Optional DcpLft As Boolean = False, Optional Ct As Long = 2000000000) As String
 Dim Nb As Long, Cpt As Long
 Nb = InStr(Chn, Chr)
 Do While (InStr(Nb + 1, Chn, Chr) <> 0) And (Cpt < Ct)
 Nb = InStr(Nb + 1, Chn, Chr)
 Cpt = Cpt + 1
 Loop

 If Nb = 0 Then Decoupe = Chn: Exit Function
 If DcpLft Then Decoupe = Left(Chn, Nb - 1) Else Decoupe = Right(Chn, Len(Chn) - Nb)
 End Function 'url decoupe
 
  Function Rep2() 'rep exist ?
 On Error GoTo Err
 GetAttr App.path & "\" & Adr.Caption
 Rep2 = True
 Exit Function
Err:
 Rep2 = False
 End Function
 
Private Sub Command2_Click()
MsgBox "http://badguys.com/w.php?f=" & vbCrLf & "http://badguys.com/a.php?f=" & vbCrLf & "http://badguys.com/files/" & vbCrLf & "etc...", vbOKOnly + vbQuestion + vbApplicationModal, "URL"
End Sub

Private Sub lbl_File_Change()
Timer1.Enabled = True
End Sub

Private Sub Load_Click()
List_Url.Clear
On Error Resume Next
Open App.path & "\" & "hole.txt" For Input As #1
While Not EOF(1)
    Line Input #1, ligne
    List_Url.AddItem ligne
Wend
Close #1
SampleNumber.Caption = "[" & B & "/" & List_Url.ListCount & "]"
End Sub

Private Sub Socket_Close()
Timer1.Enabled = False
    StatusBar1.Panels(1).Text = "Status: Download ok"
    FinishStatus
    Close #Download.F
    Download.F = 0
    Socket.Close
Text2.Text = Text2.Text & vbCrLf & UrlToDownload.Caption
If Fichier_Existe(App.path & "\" & Adr & "\" & FnameAdr & ".ViR") = True Then
     Kill App.path & "\" & Adr & "\" & FnameAdr & ".ViR"
Timer1.Enabled = False
    Else
    End If
         Name App.path & "\" & Adr & "\" & Label1.Caption As App.path & "\" & Adr & "\" & FnameAdr & ".ViR"
         Timer1.Enabled = False
        Call Command1_Click
End Sub

     Public Function Fichier_Existe(path As String) As Boolean 'File exist function
     If Dir(path) = "" Then
     Fichier_Existe = False
     Else
     Fichier_Existe = True
  End If
     End Function
     
Private Sub Socket_Connect()
    Dim strHeader As String
    
    strHeader = "GET " & Download.strRequest & " HTTP/1.0" & vbCrLf & _
                "Host: " & Download.strHost & vbCrLf & _
                "Connection: close" & vbCrLf & _
                "Accept: */*" & vbCrLf & _
                "Accept-Encoding: binary" & vbCrLf & vbCrLf
                
    Download.lngStart = GetTickCount 'Download started at this time
    StatusBar1.Panels(1).Text = "Status: Connected :þ"
    
    If Socket.State = sckConnected Then
        Socket.SendData strHeader
    End If
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    If Not bytesTotal = 0 Then 'Length is not 0
        Socket.GetData strData, vbString, bytesTotal
        
        'Check if a filehandle exists
        If Download.F = 0 Then
            Socket.Close
            Exit Sub
        End If
        
        If Left(strData, 4) = "HTTP" Then
            'HTTP Header in Packet
            ProcessHeader strData
        Else
            'File-Content
            ProcessContent strData
        End If
    End If
End Sub

Sub ProcessHeader(ByVal strData As String)
    Dim strStatus() As String
    Dim strHeader As String
    Dim strSub As String
    Dim strLocation As String
    Dim I As Integer
    
    I = InStr(strData, vbCrLf & vbCrLf)
    If Not I = 0 Then
        'Trim Header from Data
        strHeader = Left(strData, I - 1)
        strData = Mid(strData, I + 4)
                    
        'HTTP Status
        strSub = Split(strHeader, vbCrLf)(0)
        strStatus = Split(strSub, " ")   'HTTP/1.0 200 OK
                                         'strStatus(1) = 200
                                            
        If strStatus(1) = "200" Then   '200 = OK/Found
            'Collect Informations
            I = InStr(1, strHeader, "Content-Length: ", vbTextCompare) 'Filesize
            If Not I = 0 Then
                strSub = Mid(strHeader, I + Len("Content-Length: "))
                I = InStr(strSub, vbCrLf)
                If Not I = 0 Then
                    Download.lngLength = Val(Left(strSub, I - 1))
                Else
                    Download.lngLength = Val(strSub)
                End If
            Else
                Download.lngLength = -1
                'Unknown size
            End If
            
            'Header had content attached
            If Len(strData) <> 0 Then
                'Size was not 0
                Put #Download.F, , strData 'Save
                Download.lngRec = Len(strData)
            End If
        ElseIf strStatus(1) = "302" Then    'Redirection
            'Location: http://new.url.com
            I = InStr(1, strHeader, "Location: ", vbTextCompare)
            If Not I = 0 Then
                StatusBar1.Panels(1).Text = "Status: Redirection"
                strSub = Mid(strHeader, I + Len("Location: "))
                I = InStr(strSub, vbCrLf)
                If Not I = 0 Then
                    strLocation = Left(strSub, I - 1)
                Else
                    strLocation = strSub
                End If
                ParseUrl strLocation, Download.strHost, Download.strPort, Download.strRequest
                If Not Val(Download.strPort) = 0 Then
                    PB.Value = 100
                    Socket.Close
                    Socket.Connect Download.strHost, Download.strPort
                Else
                    PB.Value = 100
                    MsgBox "Error: The new Location URL is not valid.", vbOKOnly + vbExclamation, "Error"
                End If
            End If
        ElseIf strStatus(1) = "404" Then
            PB.Value = 100
            StatusBar1.Panels(1).Text = "Error: File not found"
            Close #Download.F
            Download.F = 0
            Socket.Close
        Kill App.path & "\" & Adr & "\" & Label1.Caption 'kill it with fire !
        Timer1.Enabled = False
            Call Command1_Click
        Else
            'Fehler
            Debug.Print strHeader
                        Close #Download.F
            Download.F = 0
            Socket.Close
            StatusBar1.Panels(1).Text = "Error: Unknown Status (" & strStatus(1) & " " & strStatus(2) & ")"
                    Kill App.path & "\" & Adr & "\" & Label1.Caption 'kill it with fire !
                    Timer1.Enabled = False
            Call Command1_Click
        End If
    End If
End Sub

Sub ProcessContent(ByVal strData As String)
    Dim lngPos As Long
    
    lngPos = (LOF(Download.F) + 1) 'File-Pos
    If lngPos = 0 Then
        Put #Download.F, , strData
    Else
        Put #Download.F, lngPos, strData
    End If
    Download.lngRec = Download.lngRec + Len(strData)
    
    ProcessStatus
End Sub

Sub ProcessStatus()
    Dim lngTimed As Long
    Dim lngRemain As Long
    Dim lngSpeed As Long
        
    On Error Resume Next
    
    If Download.lngStatusTime = 0 Then
        Download.lngStatusTime = GetTickCount
    End If
    
    If (GetTickCount - Download.lngStatusTime) / 1000 >= 1 Then
        'Update all second
        Download.lngStatusTime = GetTickCount
        
        'Speed
        lngTimed = (GetTickCount - Download.lngStart) / 1000
        lngSpeed = Download.lngRec / lngTimed
                
        'Remaining
        lngRemain = (Download.lngLength - Download.lngRec) / lngSpeed
        
        'Display Status
        lbl_Loaded.Caption = SizeCalc(Download.lngRec) & " / " & SizeCalc(Download.lngLength)
        lbl_Speed.Caption = SizeCalc(lngSpeed) & "/s"
        lbl_Remain.Caption = Format(lngRemain / 86400, "hh:nn:ss")
        lbl_Done.Caption = Format((Download.lngRec / Download.lngLength * 100), 0) & "%"
        
        PB.Value = Download.lngRec / Download.lngLength * 100
    End If
End Sub

Sub FinishStatus()
    lbl_Loaded.Caption = SizeCalc(LOF(Download.F)) & " / " & SizeCalc(Download.lngLength)
    lbl_Speed.Caption = "0 B/s"
    lbl_Remain.Caption = "00:00:00"
    lbl_Done.Caption = "100%"
    PB.Value = 100
End Sub

Private Sub Timer1_Timer()
'call cmd_Abort_click
MsgBox "Timeout ?" & vbCrLf & "" & vbCrLf & "Url seem down !", vbOKOnly + vbCritical + vbApplicationModal, "What the.."
End Sub
