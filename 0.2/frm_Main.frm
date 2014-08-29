VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frm_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackhole utility - Xyl2k!"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frm_Main.frx":0442
   ScaleHeight     =   5205
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "Revealer"
      Height          =   375
      Left            =   8160
      TabIndex        =   90
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   1680
      TabIndex        =   88
      Top             =   9720
      Width           =   1575
   End
   Begin VB.TextBox cdeRep2 
      Height          =   375
      Left            =   11520
      TabIndex        =   87
      Text            =   "Text16"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox Text1_iprange2 
      Height          =   285
      Left            =   11760
      TabIndex        =   86
      Text            =   "Text1_iprange2"
      Top             =   6360
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   10575
      TabIndex        =   81
      Top             =   600
      Visible         =   0   'False
      Width           =   10575
      Begin VB.Frame Frame6 
         Caption         =   "Blackhole files"
         Height          =   4215
         Left            =   120
         TabIndex        =   82
         Top             =   0
         Width           =   10335
         Begin VB.ListBox List1 
            Height          =   1425
            ItemData        =   "frm_Main.frx":1460
            Left            =   240
            List            =   "frm_Main.frx":1548
            TabIndex        =   89
            Top             =   840
            Width           =   9615
         End
         Begin VB.TextBox Text15 
            Height          =   1575
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Top             =   2400
            Width           =   9615
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Is this shit blackhole ?"
            Height          =   375
            Left            =   7680
            TabIndex        =   84
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   240
            TabIndex        =   83
            Text            =   "http://37.9.55.254/"
            Top             =   360
            Width           =   7335
         End
         Begin InetCtlsObjects.Inet Inet2 
            Left            =   9720
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
         End
      End
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   13800
      TabIndex        =   80
      Text            =   "url final"
      Top             =   6840
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4335
      ScaleWidth      =   10575
      TabIndex        =   73
      Top             =   600
      Visible         =   0   'False
      Width           =   10575
      Begin VB.CommandButton Command11 
         Caption         =   "Get Hydra"
         Height          =   375
         Left            =   240
         TabIndex        =   92
         Top             =   120
         Width           =   10095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Okay"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   3960
         Width           =   10095
      End
      Begin VB.TextBox Text12 
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   77
         Text            =   "frm_Main.frx":1D23
         Top             =   3000
         Width           =   10095
      End
      Begin VB.TextBox Text11 
         Height          =   1695
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   74
         Text            =   "frm_Main.frx":1E4D
         Top             =   840
         Width           =   10095
      End
      Begin VB.Label Label14 
         Caption         =   "Patator v0.3:"
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Python 2.7 script:"
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   5760
      TabIndex        =   57
      Text            =   $"frm_Main.frx":253B
      Top             =   11400
      Width           =   8775
   End
   Begin VB.PictureBox hydra_menu 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   10575
      TabIndex        =   54
      Top             =   600
      Visible         =   0   'False
      Width           =   10575
      Begin VB.CommandButton Command7 
         Caption         =   "Don't have Hydra ?"
         Height          =   375
         Left            =   7920
         TabIndex        =   72
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1080
         TabIndex        =   71
         Text            =   "C:\Hydra"
         Top             =   3480
         Width           =   9375
      End
      Begin VB.Frame Frame5 
         Caption         =   "Brute Force"
         Height          =   2535
         Left            =   120
         TabIndex        =   58
         Top             =   0
         Width           =   10335
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1320
            TabIndex        =   65
            Top             =   720
            Width           =   8535
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1320
            TabIndex        =   64
            Text            =   "bhadmin.php"
            Top             =   1080
            Width           =   8535
         End
         Begin VB.CheckBox Check2 
            Caption         =   "show login+pass combination for each attempt"
            Height          =   255
            Left            =   360
            TabIndex        =   63
            Top             =   2160
            Value           =   1  'Checked
            Width           =   5295
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1320
            TabIndex        =   62
            Text            =   "80"
            Top             =   1440
            Width           =   8535
         End
         Begin VB.CheckBox Check3 
            Caption         =   "waittime for responses"
            Height          =   255
            Left            =   360
            TabIndex        =   61
            Top             =   1800
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   2400
            TabIndex        =   60
            Text            =   "64"
            Top             =   1800
            Width           =   7455
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   1320
            TabIndex        =   59
            Text            =   "pwd.lst"
            Top             =   360
            Width           =   8535
         End
         Begin VB.Label Label8 
            Caption         =   "Url:"
            Height          =   255
            Left            =   360
            TabIndex        =   69
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Login page:"
            Height          =   255
            Left            =   360
            TabIndex        =   68
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Port:"
            Height          =   255
            Left            =   360
            TabIndex        =   67
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Dictionary:"
            Height          =   255
            Left            =   360
            TabIndex        =   66
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   56
         Text            =   "frm_Main.frx":25D7
         Top             =   2760
         Width           =   10335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hydra the fucker"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   3840
         Width           =   7695
      End
      Begin VB.Label Label12 
         Caption         =   "Hydra path:"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   3480
         Width           =   975
      End
   End
   Begin VB.PictureBox ip_menu 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   10575
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   10575
      Begin VB.Frame Frame4 
         Caption         =   "IP Range scanner"
         Height          =   4215
         Left            =   240
         TabIndex        =   42
         Top             =   0
         Width           =   10335
         Begin VB.CommandButton Command1_iprange 
            Caption         =   "R"
            Height          =   255
            Left            =   9720
            TabIndex        =   48
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox Text4_iprange 
            Height          =   1335
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            Top             =   2640
            Width           =   9735
         End
         Begin VB.TextBox Text3_iprange 
            Height          =   1335
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   960
            Width           =   9735
         End
         Begin VB.TextBox Text2_iprange 
            Height          =   285
            Left            =   5520
            TabIndex        =   45
            Text            =   "/bhadmin.php"
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command6_iprange 
            Caption         =   "Go !"
            Height          =   255
            Left            =   7560
            TabIndex        =   44
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox strUrl 
            Height          =   285
            Left            =   1800
            TabIndex        =   43
            Text            =   "http://146.185.238"
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label5 
            Caption         =   "."
            Height          =   255
            Left            =   4920
            TabIndex        =   53
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label3_iprange 
            Caption         =   "Timeout/403 Forbidden/404 Not Found:"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   2400
            Width           =   9735
         End
         Begin VB.Label Label2_iprange 
            Caption         =   "HTTP/1.1 200 OK:"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   720
            Width           =   9735
         End
         Begin VB.Label Label1_iprange 
            Caption         =   "0"
            Height          =   255
            Left            =   5040
            TabIndex        =   50
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Url: (http://xx.xx.xx)"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox Force_menu 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   10575
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   10575
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
         Height          =   3375
         Left            =   120
         TabIndex        =   25
         Top             =   885
         Width           =   5295
         Begin VB.CommandButton cmd_Abort 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3960
            TabIndex        =   27
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   4800
            Top             =   240
         End
         Begin VB.CheckBox af 
            Caption         =   "After download open destination folder"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   3000
            Width           =   3015
         End
         Begin MSComctlLib.ProgressBar PB 
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   2040
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lbl_CaptionFile 
            Caption         =   "File:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lbl_CaptionRate 
            Caption         =   "Speed:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lbl_CaptionLoaded 
            Caption         =   "Downloaded:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label lbl_Done 
            Alignment       =   2  'Center
            Caption         =   "0%"
            Height          =   255
            Left            =   3240
            TabIndex        =   36
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label lbl_CaptionRemain 
            Caption         =   "Remaining:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lbl_Speed 
            Caption         =   "0 B/s"
            Height          =   255
            Left            =   1800
            TabIndex        =   34
            Top             =   1320
            Width           =   3375
         End
         Begin VB.Label lbl_Remain 
            Caption         =   "00:00:00"
            Height          =   255
            Left            =   1800
            TabIndex        =   33
            Top             =   1560
            Width           =   3375
         End
         Begin VB.Label lbl_Loaded 
            Caption         =   "0 B / 0 B"
            Height          =   255
            Left            =   1800
            TabIndex        =   32
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Label lbl_File 
            Height          =   255
            Left            =   1800
            TabIndex        =   31
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label Label3 
            Caption         =   "Processing:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label UrlToDownload 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1800
            TabIndex        =   29
            Top             =   840
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Download File"
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   5295
         Begin VB.TextBox txt_Url 
            Height          =   285
            Left            =   840
            TabIndex        =   23
            Top             =   360
            Width           =   3255
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Download"
            Height          =   255
            Left            =   4200
            TabIndex        =   22
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   120
            Picture         =   "frm_Main.frx":25E5
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "Url:"
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.ListBox List_Url 
         Height          =   3375
         ItemData        =   "frm_Main.frx":2AC1
         Left            =   5520
         List            =   "frm_Main.frx":3107
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Load 
         Caption         =   "Load hole.txt"
         Height          =   255
         Left            =   5520
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "URLs found"
         Height          =   4215
         Left            =   6840
         TabIndex        =   16
         Top             =   0
         Width           =   3615
         Begin VB.TextBox Text2 
            Height          =   3735
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   360
            Width           =   3375
         End
      End
      Begin MSComDlg.CommonDialog CMD 
         Left            =   6480
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label SampleNumber 
         Alignment       =   2  'Center
         Caption         =   "[0/534]"
         Height          =   255
         Left            =   5520
         TabIndex        =   40
         Top             =   3960
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Brute Force"
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "IP Range scanner"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Force /files/"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1_iprange 
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   11760
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   11880
      Width           =   1815
   End
   Begin VB.TextBox cdeRep 
      Height          =   375
      Left            =   9120
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   10920
      Width           =   1335
   End
   Begin VB.CommandButton cmd_Download 
      Caption         =   "Download"
      Default         =   -1  'True
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      Top             =   11040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   10920
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   7800
      Top             =   10920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4950
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
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2760
      Top             =   11400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "The Truth Exists Beyond The Gate"
      Height          =   255
      Left            =   0
      TabIndex        =   91
      Top             =   2280
      Width           =   10695
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Height          =   1095
      Left            =   480
      TabIndex        =   79
      Top             =   6960
      Width           =   10095
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   10200
      Picture         =   "frm_Main.frx":3E5D
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   195
      Left            =   10200
      Picture         =   "frm_Main.frx":41FA
      Top             =   240
      Width           =   240
   End
   Begin VB.Label FnameAdr 
      Caption         =   "Label3"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   11760
      Width           =   3135
   End
   Begin VB.Label random 
      Caption         =   "Label3"
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Top             =   11640
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Label3"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   10680
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   10680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing :"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   10920
      Width           =   975
   End
   Begin VB.Label Adr 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   11280
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
Dim a, B, c, D As Integer
Dim ligne As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit

'The code is ugly and i know it.

Private Sub cdeRep2_Change()
If cdeRep2.Text = "" Then
Exit Sub
Else
If cdeRep2.Text = "HTTP/1.1 200 OK" Then
StatusBar1.Panels(1).Text = "Way to go :þ [HTTP/1.1 200 OK]: " & Text13.Text
Text15.Text = Text15.Text & vbCrLf & Text13.Text
Else
End If
Call Command9_Click
End If
End Sub

Private Sub Check2_Click()
Text3.Text = Text7.Text
If Check2.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "-V")
Else
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "")
End If

If Check3.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{TIME}", "-w " & Text8.Text)
Else
Text3.Text = Replace(Text3.Text, "{TIME}", "")
End If

Text3.Text = Replace(Text3.Text, "{DIC}", Text9.Text)
Text3.Text = Replace(Text3.Text, "{PAGE2}", Text5.Text)
Text3.Text = Replace(Text3.Text, "{URLL}", Text4.Text)
Text3.Text = Replace(Text3.Text, "{PORT}", Text6.Text)
End Sub

Private Sub Check3_Click()
Text3.Text = Text7.Text
If Check2.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "-V")
Else
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "")
End If

If Check3.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{TIME}", "-w " & Text8.Text)
Else
Text3.Text = Replace(Text3.Text, "{TIME}", "")
End If

Text3.Text = Replace(Text3.Text, "{DIC}", Text9.Text)
Text3.Text = Replace(Text3.Text, "{PAGE2}", Text5.Text)
Text3.Text = Replace(Text3.Text, "{URLL}", Text4.Text)
Text3.Text = Replace(Text3.Text, "{PORT}", Text6.Text)
End Sub

Private Sub Command10_Click()
Picture1.Visible = False
hydra_menu.Visible = False
ip_menu.Visible = False
Force_menu.Visible = False
Picture2.Visible = True
End Sub

Private Sub Command11_Click()
ShellExecute hwnd, "Open", "http://www.thc.org/thc-hydra/", "", App.path, 1
End Sub

Private Sub Command3_Click()
MsgBox "This feature is now useless since Blackhole 2.0 have come. sorry, deactivated", vbCritical, "Blackhole 2.x"
MsgBox "Anyway contact me if you have a good OCR to break captchas xylitol@malwareint.com", vbInformation, "Looking for a good VB6 OCR"
End Sub

Private Sub Command4_Click()
Force_menu.Visible = True
ip_menu.Visible = False
hydra_menu.Visible = False
Picture1.Visible = False
Picture2.Visible = False
End Sub

Private Sub Command5_Click()
ip_menu.Visible = True
Force_menu.Visible = False
hydra_menu.Visible = False
Picture1.Visible = False
Picture2.Visible = False
End Sub

Private Sub Command6_Click()
Picture1.Visible = False
hydra_menu.Visible = True
ip_menu.Visible = False
Force_menu.Visible = False
Picture2.Visible = False
End Sub



Private Sub Command7_Click()
Picture1.Visible = True
Force_menu.Visible = False
ip_menu.Visible = False
hydra_menu.Visible = True
End Sub

Private Sub Command8_Click()
Picture1.Visible = False
End Sub

Private Sub Command9_Click()
If c = List1.ListCount Then 'we have downloaded all the shit?
c = 0
Inet2.Cancel
Exit Sub
End If
D = c + 1
If c < List1.ListCount Then
Text13.Text = Text14 & List1.List(c)
List1.Selected(c) = True
c = c + 1
Else
c = 0
Exit Sub
End If
Dim hdr2, posCr2, vtData2
    Text1_iprange2.Text = ""
    On Error GoTo Probleme
   If Text14.Text = "" Then
     MsgBox "Url ?", 48, "!"
     GoTo Sortie
     End If
    
     cdeRep2.Text = ""

    Inet2.URL = Text13.Text
    
    Inet2.OpenURL
    Text1_iprange2.Text = Inet2.ResponseInfo

   
    If vtData2 = "" Then
        hdr2 = Inet2.GetHeader

        If Check4.Value = 1 Then
             Text1_iprange2.Visible = True
             Text1_iprange2.Text = hdr2
       End If
      
        posCr2 = InStr(hdr2, vbCrLf)
        hdr2 = Left(hdr2, posCr2 - 1)
        
        If Inet2.RequestTimeout > 100 And hdr2 = "" Then
                GoTo Sortie
        End If
        
    Else
        hdr2 = vtData2
    End If
    
    cdeRep2 = hdr2
   GoTo Sortie
Probleme:
   cdeRep2 = "no answer"
   StatusBar1.Panels(1).Text = "Server is dead ?"
Sortie:
End Sub

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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Inet1.Cancel
End
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
End Sub

Private Sub Image2_Click()
Unload Me
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



Private Sub Text4_Change()
Text3.Text = Text7.Text
If Check2.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "-V")
Else
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "")
End If

If Check3.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{TIME}", "-w " & Text8.Text)
Else
Text3.Text = Replace(Text3.Text, "{TIME}", "")
End If

Text3.Text = Replace(Text3.Text, "{DIC}", Text9.Text)
Text3.Text = Replace(Text3.Text, "{PAGE2}", Text5.Text)
Text3.Text = Replace(Text3.Text, "{URLL}", Text4.Text)
Text3.Text = Replace(Text3.Text, "{PORT}", Text6.Text)
End Sub

Private Sub Text5_Change()
Text3.Text = Text7.Text
If Check2.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "-V")
Else
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "")
End If

If Check3.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{TIME}", "-w " & Text8.Text)
Else
Text3.Text = Replace(Text3.Text, "{TIME}", "")
End If

Text3.Text = Replace(Text3.Text, "{DIC}", Text9.Text)
Text3.Text = Replace(Text3.Text, "{PAGE2}", Text5.Text)
Text3.Text = Replace(Text3.Text, "{URLL}", Text4.Text)
Text3.Text = Replace(Text3.Text, "{PORT}", Text6.Text)
End Sub

Private Sub Text6_Change()
Text3.Text = Text7.Text
If Check2.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "-V")
Else
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "")
End If

If Check3.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{TIME}", "-w " & Text8.Text)
Else
Text3.Text = Replace(Text3.Text, "{TIME}", "")
End If

Text3.Text = Replace(Text3.Text, "{DIC}", Text9.Text)
Text3.Text = Replace(Text3.Text, "{PAGE2}", Text5.Text)
Text3.Text = Replace(Text3.Text, "{URLL}", Text4.Text)
Text3.Text = Replace(Text3.Text, "{PORT}", Text6.Text)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer) 'only numeric
If (KeyAscii <> vbKeyBack) And Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub Text8_Change()
Text3.Text = Text7.Text
If Check2.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "-V")
Else
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "")
End If

If Check3.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{TIME}", "-w " & Text8.Text)
Else
Text3.Text = Replace(Text3.Text, "{TIME}", "")
End If

Text3.Text = Replace(Text3.Text, "{DIC}", Text9.Text)
Text3.Text = Replace(Text3.Text, "{PAGE2}", Text5.Text)
Text3.Text = Replace(Text3.Text, "{URLL}", Text4.Text)
Text3.Text = Replace(Text3.Text, "{PORT}", Text6.Text)
End Sub

Private Sub Text9_Change()
Text3.Text = Text7.Text
If Check2.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "-V")
Else
Text3.Text = Replace(Text3.Text, "{COMBINATION}", "")
End If

If Check3.Value = 1 Then
Text3.Text = Replace(Text3.Text, "{TIME}", "-w " & Text8.Text)
Else
Text3.Text = Replace(Text3.Text, "{TIME}", "")
End If

Text3.Text = Replace(Text3.Text, "{DIC}", Text9.Text)
Text3.Text = Replace(Text3.Text, "{PAGE2}", Text5.Text)
Text3.Text = Replace(Text3.Text, "{URLL}", Text4.Text)
Text3.Text = Replace(Text3.Text, "{PORT}", Text6.Text)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer) 'only numeric
If (KeyAscii <> vbKeyBack) And Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
'call cmd_Abort_click
MsgBox "Timeout ?" & vbCrLf & "" & vbCrLf & "Url seem down !", vbOKOnly + vbCritical + vbApplicationModal, "What the.."
End Sub

'ip range scann code
Private Sub cdeRep_Change()
If cdeRep.Text = "HTTP/1.1 200 OK" Then
StatusBar1.Panels(1).Text = "HTTP/1.1 200 OK :)"
Text3_iprange.Text = Text3_iprange.Text & vbCrLf & strUrl.Text & "." & Label1_iprange.Caption & Text2_iprange.Text
If Label1_iprange.Caption = 255 Then
Exit Sub
Else
Label1_iprange.Caption = Label1_iprange.Caption + 1
Call Command6_iprange_Click
End If
Else
End If
If cdeRep.Text = "Le serveur ne répond pas" Then
StatusBar1.Panels(1).Text = "Timeout ?"
Text4_iprange.Text = Text4_iprange.Text & vbCrLf & strUrl.Text & "." & Label1_iprange.Caption & Text2_iprange.Text & " => Time Out !"
If Label1_iprange.Caption = 255 Then
Exit Sub
Else
Label1_iprange.Caption = Label1_iprange.Caption + 1
Call Command6_iprange_Click
End If
End If
If cdeRep.Text = "HTTP/1.1 403 Forbidden" Then
StatusBar1.Panels(1).Text = "HTTP/1.1 403 Forbidden :("
Text4_iprange.Text = Text4_iprange.Text & vbCrLf & strUrl.Text & "." & Label1_iprange.Caption & Text2_iprange.Text & " => HTTP/1.1 403 Forbidden"
If Label1_iprange.Caption = 255 Then
Exit Sub
Else
Label1_iprange.Caption = Label1_iprange.Caption + 1
Call Command6_iprange_Click
End If
End If
If cdeRep.Text = "HTTP/1.1 404 Not Found" Then
StatusBar1.Panels(1).Text = "HTTP/1.1 404 Not Found :("
Text4_iprange.Text = Text4_iprange.Text & vbCrLf & strUrl.Text & "." & Label1_iprange.Caption & Text2_iprange.Text & " => HTTP/1.1 404 Not Found"
If Label1_iprange.Caption = 255 Then
Exit Sub
Else
Label1_iprange.Caption = Label1_iprange.Caption + 1
Call Command6_iprange_Click
End If
End If
End Sub

Private Sub Command6_iprange_Click() 'Test URLs
Dim hdr, posCr, vtData
    Text1_iprange.Text = ""
    On Error GoTo Probleme
   If strUrl.Text = "" Then 'strUrl.Text contientl'URL
     MsgBox "Url ?", 48, "!"
     GoTo Sortie
     End If
    
     cdeRep.Text = "" 'On efface la case réponse

    Inet1.URL = strUrl & "." & Label1_iprange.Caption & Text2_iprange.Text 'On lui dit l'URL à ouvrir
    
    Inet1.OpenURL 'Il reçoit l'ordre d'ouvrir
    Text1_iprange.Text = Inet1.ResponseInfo

     'On traite la réponse éventuelle
    If vtData = "" Then
        hdr = Inet1.GetHeader

        If Check1.Value = 1 Then
             Text1_iprange.Visible = True
             Text1_iprange.Text = hdr
       End If
      
        posCr = InStr(hdr, vbCrLf)
        hdr = Left(hdr, posCr - 1)
        
        If Inet1.RequestTimeout > 100 And hdr = "" Then
                GoTo Sortie
        End If
        
    Else
        hdr = vtData
    End If
    
    cdeRep = hdr
   GoTo Sortie
Probleme:
   cdeRep = "Le serveur ne répond pas"
Sortie:
End Sub

Private Sub Command1_iprange_Click()
Label1_iprange.Caption = "0"
End Sub
