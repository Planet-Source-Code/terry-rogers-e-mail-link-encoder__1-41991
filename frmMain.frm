VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML E-Mail Link Encoder"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPreview 
      Caption         =   " Preview Panel "
      Height          =   1935
      Left            =   158
      TabIndex        =   10
      Top             =   4515
      Width           =   8115
      Begin VB.TextBox txtStatus 
         Height          =   360
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1470
         Width           =   7785
      End
      Begin SHDocVwCtl.WebBrowser Browser 
         Height          =   1065
         Left            =   150
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   330
         Width           =   7785
         ExtentX         =   13732
         ExtentY         =   1879
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Frame frmMain 
      Caption         =   " Encoded Results "
      Height          =   1545
      Left            =   158
      TabIndex        =   9
      Top             =   2895
      Width           =   8115
      Begin VB.TextBox txtResults 
         Height          =   1050
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   330
         Width           =   7785
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   405
      Left            =   4448
      TabIndex        =   8
      Top             =   2265
      Width           =   1995
   End
   Begin VB.CommandButton cmdEncoder 
      Caption         =   "Encode"
      Height          =   405
      Left            =   1988
      TabIndex        =   7
      Top             =   2265
      Width           =   1995
   End
   Begin VB.TextBox txtText 
      Height          =   360
      Left            =   1733
      TabIndex        =   4
      Top             =   1140
      Width           =   6465
   End
   Begin VB.TextBox txtStatusBarText 
      Height          =   360
      Left            =   1733
      TabIndex        =   6
      Top             =   1665
      Width           =   6465
   End
   Begin VB.TextBox txtEmailAddress 
      Height          =   360
      Left            =   1733
      TabIndex        =   2
      Top             =   615
      Width           =   6465
   End
   Begin VB.Label lbTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "HTML Link Encoder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3308
      TabIndex        =   0
      Top             =   165
      Width           =   1815
   End
   Begin VB.Label lbText 
      AutoSize        =   -1  'True
      Caption         =   "Text:"
      Height          =   240
      Left            =   203
      TabIndex        =   3
      Top             =   1245
      Width           =   450
   End
   Begin VB.Label lbStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status Bar Text:"
      Height          =   240
      Left            =   233
      TabIndex        =   5
      Top             =   1725
      Width           =   1395
   End
   Begin VB.Label lbEmailAddress 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail Address:"
      Height          =   240
      Left            =   233
      TabIndex        =   1
      Top             =   675
      Width           =   1350
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Author           : Terry Rogers
' * Web Site         : http://www.terryrogers.net
' * E-Mail           : terry@terryrogers.net
' * Date             : 12/30/2002
' * Time             : 08:58
' * Module Name      : frmMain
' * Module Filename  : frmMain.frm
' **********************************************************************

Option Explicit

Public vEmailAddress, vText, vStatusBarText As String

Function EncodeHTML(vHTML As String)
    ' #VBIDEUtils#************************************************************
    ' * Author           : Terry Rogers
    ' * Web Site         : http://www.terryrogers.net
    ' * E-Mail           : terry@terryrogers.net
    ' * Date             : 12/30/2002
    ' * Time             : 08:58
    ' * Module Name      : frmMain
    ' * Module Filename  : frmMain.frm
    ' * Procedure Name   : EncodeHTML
    ' * Parameters       :
    ' *                    vHTML As String
    ' **********************************************************************
    Dim k                As Integer
    For k = 1 To Len(vHTML)
        EncodeHTML = EncodeHTML & "&#" & Asc(Mid(vHTML, k, 1)) & ";"
    Next k
End Function

Private Sub Browser_StatusTextChange(ByVal Text As String)
    ' #VBIDEUtils#************************************************************
    ' * Author           : Terry Rogers
    ' * Web Site         : http://www.terryrogers.net
    ' * E-Mail           : terry@terryrogers.net
    ' * Date             : 12/30/2002
    ' * Time             : 08:58
    ' * Module Name      : frmMain
    ' * Module Filename  : frmMain.frm
    ' * Procedure Name   : Browser_StatusTextChange
    ' * Parameters       :
    ' *                    ByVal Text As String
    ' **********************************************************************
    txtStatus.Text = Text
End Sub

Private Sub cmdEncoder_Click()
    ' #VBIDEUtils#************************************************************
    ' * Author           : Terry Rogers
    ' * Web Site         : http://www.terryrogers.net
    ' * E-Mail           : terry@terryrogers.net
    ' * Date             : 12/30/2002
    ' * Time             : 08:58
    ' * Module Name      : frmMain
    ' * Module Filename  : frmMain.frm
    ' * Procedure Name   : cmdEncoder_Click
    ' * Parameters       :
    ' **********************************************************************
    If txtEmailAddress.Text = "" Or IsNull(txtEmailAddress.Text) Then
        Call MsgBox("Please enter an email address.", vbExclamation, "E-Mail Link Encoder")
        txtEmailAddress.SetFocus
        Exit Sub
    End If
    If txtText.Text = "" Or IsNull(txtText.Text) Then
        Call MsgBox("Please enter the text you wish to be displayed.", vbExclamation, "E-Mail Link Encoder")
        txtText.SetFocus
        Exit Sub
    End If
    If txtStatusBarText.Text = "" Or IsNull(txtStatusBarText.Text) Then
        vEmailAddress = EncodeHTML("mailto:" & txtEmailAddress.Text)
        vText = EncodeHTML(txtText.Text)
        txtResults.Text = "<a href=" & Chr(34) & vEmailAddress & Chr(34) & _
            ">" & vText & "</a>"
        If FileExists("C:\LinkEncoder.dat") = True Then Kill "C:\LinkEncoder.dat"
        Open "C:\LinkEncoder.dat" For Output As #1
        Print #1, "<html><head><title>Link Encoder</title></head>"
        Print #1, "<body bgcolor=" & Chr(13) & "#FFFFFF" & Chr(13) & " text=" & Chr(13) & "#0000FF" & Chr(13) & " link=" & Chr(13) & "#0000FF" & Chr(13) & " vlink=" & Chr(13) & "#0000FF" & Chr(13) & " alink=" & Chr(13) & "#0000FF" & Chr(13) & ">"
        Print #1, "<center>"
        Print #1, txtResults.Text
        Print #1, "</center>"
        Print #1, "</body></html>"
        Close #1
        Browser.Navigate "C:\LinkEncoder.dat"
    Else
        vEmailAddress = EncodeHTML("mailto:" & txtEmailAddress.Text)
        vText = EncodeHTML(txtText.Text)
        vStatusBarText = EncodeHTML(txtStatusBarText.Text)
        txtResults.Text = "<a href=" & Chr(34) & vEmailAddress & Chr(34) & _
            "onMouseOver=" & Chr(34) & "self.status='" & vStatusBarText & _
            "';return true" & Chr(34) & "onMouseOut=" & Chr(34) & _
            "self.status='';return true" & Chr(34) & ">" & vText & "</a>"
        If FileExists("C:\LinkEncoder.dat") = True Then Kill "C:\LinkEncoder.dat"
        Open "C:\LinkEncoder.dat" For Output As #1
        Print #1, "<html><head><title>Link Encoder</title></head>"
        Print #1, "<body bgcolor=" & Chr(13) & "#FFFFFF" & Chr(13) & " text=" & Chr(13) & "#0000FF" & Chr(13) & " link=" & Chr(13) & "#0000FF" & Chr(13) & " vlink=" & Chr(13) & "#0000FF" & Chr(13) & " alink=" & Chr(13) & "#0000FF" & Chr(13) & ">"
        Print #1, "<center>"
        Print #1, txtResults.Text
        Print #1, "</center>"
        Print #1, "</body></html>"
        Close #1
        Browser.Navigate "C:\LinkEncoder.dat"
    End If
End Sub

Private Sub cmdQuit_Click()
    ' #VBIDEUtils#************************************************************
    ' * Author           : Terry Rogers
    ' * Web Site         : http://www.terryrogers.net
    ' * E-Mail           : terry@terryrogers.net
    ' * Date             : 12/30/2002
    ' * Time             : 08:58
    ' * Module Name      : frmMain
    ' * Module Filename  : frmMain.frm
    ' * Procedure Name   : cmdQuit_Click
    ' * Parameters       :
    ' **********************************************************************
    If FileExists("C:\LinkEncoder.dat") = True Then Kill "C:\LinkEncoder.dat"
    Unload Me
End Sub

Private Sub Form_Load()
    ' #VBIDEUtils#************************************************************
    ' * Author           : Terry Rogers
    ' * Web Site         : http://www.terryrogers.net
    ' * E-Mail           : terry@terryrogers.net
    ' * Date             : 12/30/2002
    ' * Time             : 08:58
    ' * Module Name      : frmMain
    ' * Module Filename  : frmMain.frm
    ' * Procedure Name   : Form_Load
    ' * Parameters       :
    ' **********************************************************************
    If App.PrevInstance Then
        ActivatePrevInstance
    End If
    If FileExists("C:\LinkEncoder.dat") = True Then Kill "C:\LinkEncoder.dat"
    Open "C:\LinkEncoder.dat" For Output As #1
    Print #1,
    Close #1
    Browser.Navigate "C:\LinkEncoder.dat"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' #VBIDEUtils#************************************************************
    ' * Author           : Terry Rogers
    ' * Web Site         : http://www.terryrogers.net
    ' * E-Mail           : terry@terryrogers.net
    ' * Date             : 12/30/2002
    ' * Time             : 08:58
    ' * Module Name      : frmMain
    ' * Module Filename  : frmMain.frm
    ' * Procedure Name   : Form_Unload
    ' * Parameters       :
    ' *                    Cancel As Integer
    ' **********************************************************************
    End
End Sub

Function FileExists%(filename$)
    ' #VBIDEUtils#************************************************************
    ' * Author           : Terry Rogers
    ' * Web Site         : http://www.terryrogers.net
    ' * E-Mail           : terry@terryrogers.net
    ' * Date             : 12/30/2002
    ' * Time             : 08:58
    ' * Module Name      : frmMain
    ' * Module Filename  : frmMain.frm
    ' * Procedure Name   : FileExists
    ' * Parameters       :
    ' *                    filename$
    ' **********************************************************************
    Dim f%
    On Error Resume Next
    f% = FreeFile
    Open filename$ For Input As #f%
    Close #f%
    FileExists% = Not (Err <> 0)
End Function

