VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "控制台 ControlCentre"
   ClientHeight    =   7095
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "运行"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3840
      Top             =   2520
   End
   Begin VB.CommandButton Command3 
      Caption         =   "论证安全性"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "指定目录读取文件"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3000
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "发送文本"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "起飞"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "请输人文本"
      Top             =   4920
      Width           =   12015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Cascadia Code SemiLight"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   12015
   End
   Begin VB.Menu 管理 
      Caption         =   "Key"
      Begin VB.Menu keyman 
         Caption         =   "管理"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim httpSec As Object
Dim http As Object

Private Sub Command2_Click()
Form3.Show

End Sub

Private Sub keyman_Click()
Form2.Show

End Sub


'==================== 窗体加载 ====================
Private Sub Form_Load()
    Call LoadApiKey
    
    Timer1.Interval = 500
    Timer1.Enabled = False
    
    Command1.Caption = "发送"
    Text1.Text = ""
    Text2.Text = ""
End Sub

'==================== 读取 Key ====================
Private Sub LoadApiKey()
    Dim strFile As String
    Dim iFile As Integer
    Dim keyText As String
    Dim lineText As String
    
    strFile = "C:\aidata.ky"
    
    If Dir(strFile) = "" Then
        Form2.Label2.Caption = ""
        Exit Sub
    End If
    
    iFile = FreeFile
    Open strFile For Input As #iFile
    
    Do Until EOF(iFile)
        Line Input #iFile, lineText
        keyText = keyText & lineText
    Loop
    Close #iFile
    
    Form2.Label2.Caption = Trim$(keyText)
End Sub

'==================== 发送按钮 ====================
Private Sub Command1_Click()
    Dim url As String
    Dim apiKey As String
    Dim postData As String
    Dim prompt As String
    
    ' 每次发送前重新读 key
    Call LoadApiKey
    apiKey = Trim$(Form2.Label2.Caption)
    
    If apiKey = "" Then
        MsgBox "API Key 为空，请检查 C:\aidata.ky 文件！", vbCritical
        Exit Sub
    End If
    
    prompt = Trim$(Text2.Text)
    If prompt = "" Then
        MsgBox "请先在 Text2 输入内容！", vbExclamation
        Exit Sub
    End If
    
    ' ★ 核心改动：强制前缀，让 AI 只吐代码
    prompt = "你只能输出cmd纯代码，禁止任何解释、禁止 markdown 正文说明、禁止开场白和结束语。如果用户需求不明确，输出空字符串。以下是需求：" & vbCrLf & prompt
    
    ' JSON 转义
    prompt = Replace(prompt, "\", "\\")
    prompt = Replace(prompt, """", "\""")
    prompt = Replace(prompt, vbCrLf, "\n")
    prompt = Replace(prompt, vbLf, "\n")
    
    postData = "{""model"":""deepseek-chat"",""messages"":[{""role"":""user"",""content"":""" & prompt & """}],""max_tokens"":2048,""temperature"":0.1}"
    
    url = "https://api.deepseek.com/v1/chat/completions"
    
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set http = CreateObject("MSXML2.ServerXMLHTTP.3.0")
    End If
    If http Is Nothing Then
        MsgBox "? 创建 HTTP 对象失败！", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    http.setTimeouts 5000, 10000, 10000, 30000
    http.Open "POST", url, True
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.Send postData
    
    Command1.Enabled = False
    Command1.Caption = "请求中..."
    Text1.Text = "正在等待 AI 回复..."
    Timer1.Enabled = True
End Sub

'==================== Timer 轮询 ====================
Private Sub Timer1_Timer()
    If http Is Nothing Then
        Timer1.Enabled = False
        Exit Sub
    End If
    
    If http.readyState <> 4 Then Exit Sub
    
    Timer1.Enabled = False
    
    If http.Status = 200 Then
        Text1.Text = ExtractReply(http.responseText)
    Else
        Text1.Text = "请求失败 (HTTP " & http.Status & ")" & vbCrLf & vbCrLf & http.responseText
    End If
    
    Command1.Enabled = True
    Command1.Caption = "发送"
    Set http = Nothing
End Sub

'==================== 提取回复 ====================
Private Function ExtractReply(ByVal json As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim content As String
    Dim i As Long
    
    startPos = InStr(json, """content"":""")
    If startPos = 0 Then
        ExtractReply = json
        Exit Function
    End If
    
    startPos = startPos + 11
    
    For i = startPos To Len(json)
        If Mid$(json, i, 1) = """" Then
            If i > 1 Then
                If Mid$(json, i - 1, 1) <> "\" Then
                    endPos = i - 1
                    Exit For
                End If
            Else
                endPos = i - 1
                Exit For
            End If
        End If
    Next i
    
    If endPos >= startPos Then
        content = Mid$(json, startPos, endPos - startPos + 1)
        content = Replace(content, "\\", "\")
        content = Replace(content, "\""", """")
        content = Replace(content, "\n", vbCrLf)
        content = Replace(content, "\t", vbTab)
        ExtractReply = content
    Else
        ExtractReply = json
    End If
End Function

'==================== 清理 ====================
Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Set http = Nothing
End Sub



 ' ← 安全性检测用的独立 HTTP 对象，不能跟聊天共用

'==================== Command3 —— 安全性检测 ====================
Private Sub Command3_Click()
    Dim url As String
    Dim apiKey As String
    Dim postData As String
    Dim codeText As String
    Dim safePrompt As String
    
    ' 读 key
    Call LoadApiKey
    apiKey = Trim$(Form2.Label2.Caption)
    
    If apiKey = "" Then
        MsgBox "API Key 为空，请检查 C:\aidata.ky 文件！", vbCritical
        Exit Sub
    End If
    
    ' 读 Text1 的代码
    codeText = Trim$(Text1.Text)
    If codeText = "" Then
        MsgBox "Text1 中没有代码可检测！", vbExclamation
        Exit Sub
    End If
    
    ' ★ 构造专用 prompt —— 要求分析功能和安全性
    safePrompt = "请分析以下cmd/bat代码的功能和安全性。" & _
                 "格式要求：先说【功能概述】，再说【安全性评估】（标注是否有恶意行为、破坏性操作、信息泄露风险等），最后给【风险等级：安全/低风险/中风险/高风险】。" & _
                 "如果用户给的不是代码或为空，输出空字符串。" & _
                 "禁止输出与代码无关的内容。" & vbCrLf & vbCrLf & codeText
    
    ' JSON 转义
    safePrompt = Replace(safePrompt, "\", "\\")
    safePrompt = Replace(safePrompt, """", "\""")
    safePrompt = Replace(safePrompt, vbCrLf, "\n")
    safePrompt = Replace(safePrompt, vbLf, "\n")
    
    postData = "{""model"":""deepseek-chat"",""messages"":[{""role"":""user"",""content"":""" & safePrompt & """}],""max_tokens"":2048,""temperature"":0.1}"
    
    url = "https://api.deepseek.com/v1/chat/completions"
    
    On Error Resume Next
    Set httpSec = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set httpSec = CreateObject("MSXML2.ServerXMLHTTP.3.0")
    End If
    If httpSec Is Nothing Then
        MsgBox "创建 HTTP 对象失败！", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    httpSec.setTimeouts 5000, 10000, 10000, 30000
    httpSec.Open "POST", url, True
    httpSec.setRequestHeader "Content-Type", "application/json"
    httpSec.setRequestHeader "Authorization", "Bearer " & apiKey
    httpSec.Send postData
    
    Command3.Enabled = False
    Command3.Caption = "分析中..."
    Timer2.Enabled = True
End Sub

'==================== Timer2 —— 轮询安全性检测结果 ====================
Private Sub Timer2_Timer()
    If httpSec Is Nothing Then
        Timer2.Enabled = False
        Exit Sub
    End If
    
    If httpSec.readyState <> 4 Then Exit Sub
    
    Timer2.Enabled = False
    
    Dim reply As String
    
    If httpSec.Status = 200 Then
        reply = ExtractReply(httpSec.responseText)
        If Trim$(reply) = "" Then
            MsgBox "AI 未返回分析结果（可能输入的不是有效代码）。", vbInformation
        Else
            ' ★ 用 MsgBox 弹出来
            MsgBox reply, vbInformation + vbOKOnly, "?? 代码安全性分析报告"
        End If
    Else
        MsgBox "请求失败 (HTTP " & httpSec.Status & ")" & vbCrLf & vbCrLf & httpSec.responseText, vbCritical
    End If
    
    Command3.Enabled = True
    Command3.Caption = "安全检测"
    Set httpSec = Nothing
End Sub

