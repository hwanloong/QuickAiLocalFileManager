VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文件"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9600
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "读取并放到输入框"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   2880
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2160
      Width           =   4935
   End
   Begin VB.FileListBox File1 
      Height          =   2070
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2610
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1215
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive

End Sub


Option Explicit
Private Sub Timer1_Timer()
Label1.Caption = File1.path


End Sub
Private Sub Command1_Click()
    Dim strPath As String
    Dim strFile As String
    Dim strResult As String
    
    ' 从 Label1 取目录
    strPath = Label1.Caption
    
    ' 确保目录末尾有 \
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    
    ' 查找第一个文件
    strFile = Dir(strPath & "*.*")
    
    ' 循环读取所有文件
    Do While strFile <> ""
        strResult = strResult & strPath & strFile & vbCrLf
        strFile = Dir
    Loop
    
    ' 显示到 Text1
    Text1.Text = strResult
    Form1.Text2 = Text1.Text
End Sub
