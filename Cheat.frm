VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form forMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "在线课程作弊器"
   ClientHeight    =   10005
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   9810
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSelect 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5000
      Left            =   4995
      TabIndex        =   4
      Top             =   0
      Width           =   5000
      Begin VB.ListBox lisProjects 
         Height          =   3300
         Left            =   500
         TabIndex        =   6
         Top             =   1000
         Width           =   4020
      End
      Begin VB.Label Label2 
         Caption         =   "请登陆:"
         Height          =   240
         Left            =   270
         TabIndex        =   5
         Top             =   270
         Width           =   1860
      End
   End
   Begin VB.Frame fraLogin 
      BorderStyle     =   0  'None
      Height          =   5000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5000
      Begin VB.Frame fraWeb 
         BorderStyle     =   0  'None
         Height          =   3345
         Left            =   500
         TabIndex        =   2
         Top             =   1000
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
         Caption         =   "请登陆:"
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

Dim prjName() As String
Dim prjUrl() As String

Private Sub Form_Click()
    addTime
End Sub

Private Sub Form_Load()
    initLogin
End Sub

Private Function initLogin()
    web.Navigate "http://www.0755tt.com/"
End Function

Private Sub lisProjects_DblClick()
    selectProject
End Sub

Private Sub lisProjects_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        selectProject
    End If
End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If web.LocationURL = "http://www.0755tt.com/" Then
        If web.ReadyState = READYSTATE_COMPLETE Then
            Debug.Print "登陆准备完毕"
        End If
    ElseIf web.LocationURL = "http://www.0755tt.com/index!index" Then
        If web.ReadyState = READYSTATE_COMPLETE Then
            web.Navigate "http://www.0755tt.com/myCourse!myStudyCourseList"
        End If
    ElseIf web.LocationURL = "http://www.0755tt.com/myCourse!myStudyCourseList" Then
        If web.ReadyState = READYSTATE_COMPLETE Then
            web.Document.parentWindow.execScript "var funTitle = function () { var objs = $('.content h2 a strong'); var strTitles = ''; for (var i = 0; i < objs.length; i++) { strTitles += (strTitles ? ',' : '') + $(objs[i]).text(); } return strTitles; }; var funValue = function () { var objs = $('.content .content .styled a'); var strUrls = ''; for (var i = 0; i < objs.length; i++) { var courseID = $(objs[i]).attr('href').replace(/^.*courseID=(.*)$/, '$1'); strUrls += (strUrls ? ',' : '') + courseID; } return strUrls; }; var pakReturns = funTitle() + '|' + funValue();"
            parseProjectListPage web.Document.Script.pakReturns
            
            'Text1.text = web.Document.body.innerhtml
            'parseProjectListPage web.Document.body.innerhtml
        End If
    End If
End Sub

Private Sub parseProjectListPage(ByVal text As String)
    Dim arrStr() As String
    Dim i As Integer
    arrStr = Split(text, "|")
    prjName = Split(arrStr(0), ",")
    prjUrl = Split(arrStr(1), ",")
    For i = 0 To UBound(prjName)
        lisProjects.AddItem prjName(i) + " (" + prjUrl(i) + ")"
    Next i
End Sub

Private Sub selectProject()
    MsgBox lisProjects.ListIndex
    On Error Resume Next
    MsgBox prjUrl(lisProjects.ListIndex)
End Sub

Private Sub addTime()
    On Error Resume Next
    Dim time As Integer
    time = 0
    time = InputBox("请输入需要增加的小时数:")
    If time = 0 Then
        MsgBox "输入无效,请重试."
        Return
    End If
End Sub
