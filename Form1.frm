VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "xFinger v1.0"
   ClientHeight    =   2640
   ClientLeft      =   5685
   ClientTop       =   4830
   ClientWidth     =   5940
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5940
   Begin VB.CommandButton Find 
      Caption         =   "&Find "
      Default         =   -1  'True
      Height          =   345
      Left            =   540
      TabIndex        =   0
      Top             =   2070
      Width           =   1245
   End
   Begin VB.TextBox Uname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   165
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin MSComctlLib.ListView wList 
      Height          =   1545
      Left            =   2295
      TabIndex        =   3
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2725
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Workstation Name"
         Object.Width           =   5997
      EndProperty
   End
   Begin VB.TextBox Dname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   165
      TabIndex        =   1
      Top             =   855
      Width           =   1935
   End
   Begin VB.Frame wFinger 
      Height          =   2505
      Left            =   45
      TabIndex        =   4
      Top             =   90
      Width           =   5850
      Begin VB.Label Status 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   2250
         TabIndex        =   7
         Top             =   270
         Width           =   3465
      End
      Begin VB.Label Label2 
         Caption         =   "User Name:"
         Height          =   285
         Left            =   105
         TabIndex        =   6
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Domain:"
         Height          =   285
         Left            =   105
         TabIndex        =   5
         Top             =   450
         Width           =   1605
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Find_Click()
    Dim User            As IADsUser
    Dim textUser        As String
    Dim textDomain      As String
    Dim userFullName    As String
    Dim userHome        As String
    Dim fService        As IADsFileServiceOperations
    Dim Session         As IADsSession
    On Error Resume Next
    DoEvents
    wList.ListItems.Clear
    Status.Caption = ""
    textDomain = Trim(Dname.Text)
    textUser = Trim(Uname.Text)
    If textDomain = "" Or textUser = "" Then
        Status.Caption = "Domain or User Name is not specified!"
        Exit Sub
    End If
    Set User = GetObject("WinNT://" & textDomain & "/" & textUser & ",user")
    If Err.Number <> 0 Then Status.Caption = "Domain or User does not exist.": Exit Sub
    userFullName = User.FullName
    userHome = User.HomeDirectory
    If userHome = "" Then Status.Caption = "Home Server for user ''" & _
                         userFullName & "'' is not specified!": Exit Sub
    sName = Split(userHome, "\")
    Set fService = GetObject("WinNT://" & sName(2) & "/LanmanServer")
    If Err.Number <> 0 Then Status.Caption = "Can not connect to user Home Server. Server is down, " & _
                            "or you do not have enough privileges!": Exit Sub
    For Each Session In fService.Sessions
        If LCase(Session.User) = LCase(textUser) Then
             Set ItemAdd = wList.ListItems.Add(, , Session.Computer)
        End If
    Next
    If wList.ListItems.Count = 0 Then Status.Caption = "User ''" & _
                                     userFullName & "'' is not logged on.": Exit Sub
    Status.Caption = "User ''" & userFullName & "'' logged on to:"
End Sub
