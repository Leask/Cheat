VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form forMain 
   Caption         =   "ÇëµÇÂ½"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16920
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   16920
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame fraLogin 
      Height          =   4560
      Left            =   900
      TabIndex        =   0
      Top             =   270
      Width           =   6360
      Begin SHDocVwCtl.WebBrowser web 
         Height          =   3480
         Left            =   990
         TabIndex        =   2
         Top             =   810
         Width           =   4380
         ExtentX         =   7726
         ExtentY         =   6138
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
      Begin VB.Label Label1 
         Caption         =   "ÇëµÇÂ½"
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   180
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

