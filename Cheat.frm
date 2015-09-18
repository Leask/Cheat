VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form forMain 
   Caption         =   "ÇëµÇÂ½"
   ClientHeight    =   10005
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   10395
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame fraLogin 
      Height          =   6810
      Left            =   900
      TabIndex        =   0
      Top             =   270
      Width           =   9420
      Begin VB.Frame fraWeb 
         Height          =   3345
         Left            =   315
         TabIndex        =   2
         Top             =   1170
         Width           =   4020
         Begin SHDocVwCtl.WebBrowser web 
            Height          =   19995
            Left            =   -13365
            TabIndex        =   3
            Top             =   -2970
            Width           =   19995
            ExtentX         =   35278
            ExtentY         =   35278
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
            Location        =   ""
         End
      End
      Begin VB.Label Label1 
         Caption         =   "ÇëµÇÂ½:"
         Height          =   240
         Left            =   270
         TabIndex        =   1
         Top             =   270
         Width           =   1860
      End
   End
End
Attribute VB_Name = "forMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    initLogin
End Sub

Private Function initLogin()
    web.Navigate "http://www.0755tt.com/"
End Function

Private Sub web_StatusTextChange(ByVal Text As String)
    If web.LocationURL = "http://www.0755tt.com/" Then
        If web.ReadyState = READYSTATE_COMPLETE Then
            Debug.Print "µÇÂ½×¼±¸Íê±Ï"
        End If
    ElseIf web.LocationURL = "http://www.0755tt.com/index!index" Then
        If web.ReadyState = READYSTATE_COMPLETE Then
            web.Navigate "http://www.0755tt.com/myCourse!myStudyCourseList"
        End If
    End If
End Sub
