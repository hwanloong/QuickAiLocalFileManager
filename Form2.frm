VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My API key"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6135
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4680
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   1815
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
      Height          =   420
      Left            =   360
      TabIndex        =   0
      Text            =   "sk-"
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5040
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "注意，请妥善保管Key。Key会通过编码进行保存。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim strFile As String
    Dim iFile As Integer
    
    ' 目标文件路径
    strFile = "C:\aidata.ky"
    
    ' 获取一个可用的文件号
    iFile = FreeFile
    
    ' 判断文件是否存在
    If Dir(strFile) = "" Then
        ' 文件不存在，创建并写入
        Open strFile For Output As #iFile
        Print #iFile, Text1.Text
        Close #iFile
        MsgBox "完成", vbInformation
    Else
        ' 文件存在，覆盖写入
        Open strFile For Output As #iFile
        Print #iFile, Text1.Text
        Close #iFile
        MsgBox "完成", vbInformation
    End If
End Sub


' Timer1 定时读取 .ky 文件内容显示到 Label2
Private Sub Timer1_Timer()
    Dim strFile As String
    Dim iFile As Integer
    Dim strContent As String
    Dim lineText As String
    
    strFile = "C:\aidata.ky"
    
    ' 判断文件是否存在
    If Dir(strFile) = "" Then
        Label2.Caption = "文件不存在"
        Exit Sub
    End If
    
    ' 读取文件内容
    iFile = FreeFile
    Open strFile For Input As #iFile
    
    Do Until EOF(iFile)
        Line Input #iFile, lineText
        strContent = strContent & lineText & vbCrLf
    Loop
    
    Close #iFile
    
    ' 显示到 Label2（去掉末尾多余换行）
    Label2.Caption = Left$(strContent, Len(strContent) - 2)
End Sub
