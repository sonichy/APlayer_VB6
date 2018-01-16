VERSION 5.00
Object = "{A9757030-96F6-485E-A8AB-5B5137462472}#1.0#0"; "APlayerUI_1.5.0.26.dll"
Begin VB.Form Form1 
   Caption         =   "海天鹰播放器"
   ClientHeight    =   5205
   ClientLeft      =   165
   ClientTop       =   540
   ClientWidth     =   8325
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5205
   ScaleWidth      =   8325
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4170
      Left            =   6800
      TabIndex        =   1
      Top             =   0
      Width           =   1500
   End
   Begin APlayerUILibCtl.Player Player1 
      Height          =   5175
      Left            =   0
      OleObjectBlob   =   "Form1.frx":1CCA
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.Menu file 
      Caption         =   "文件"
      Begin VB.Menu open 
         Caption         =   "打开"
         Shortcut        =   ^O
      End
      Begin VB.Menu openURL 
         Caption         =   "打开网址"
         Shortcut        =   ^U
      End
      Begin VB.Menu live 
         Caption         =   "直播"
         Shortcut        =   ^L
      End
      Begin VB.Menu quit 
         Caption         =   "退出"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu video 
      Caption         =   "视频"
      Begin VB.Menu fullscreen 
         Caption         =   "全屏"
         Shortcut        =   ^F
      End
      Begin VB.Menu info 
         Caption         =   "信息"
         Shortcut        =   ^I
      End
      Begin VB.Menu capture 
         Caption         =   "截图"
         Shortcut        =   ^C
      End
      Begin VB.Menu pathCapture 
         Caption         =   "截图目录"
         Shortcut        =   ^P
      End
      Begin VB.Menu forward 
         Caption         =   "快进"
         Shortcut        =   ^N
      End
      Begin VB.Menu backward 
         Caption         =   "快退"
         Shortcut        =   ^B
      End
      Begin VB.Menu subtitle 
         Caption         =   "字幕"
         Begin VB.Menu chooseSubtitle 
            Caption         =   "选择字幕"
         End
         Begin VB.Menu switchSubtitle 
            Caption         =   "字幕开关"
         End
      End
      Begin VB.Menu GIF 
         Caption         =   "生成GIF"
      End
      Begin VB.Menu flipH 
         Caption         =   "水平翻转"
      End
      Begin VB.Menu flipV 
         Caption         =   "垂直翻转"
      End
      Begin VB.Menu rotate 
         Caption         =   "旋转"
      End
   End
   Begin VB.Menu sound 
      Caption         =   "音频"
      Begin VB.Menu volumeUp 
         Caption         =   "增加音量"
         Shortcut        =   ^X
      End
      Begin VB.Menu volumeDown 
         Caption         =   "减小音量"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mute 
         Caption         =   "静音"
         Shortcut        =   ^M
      End
      Begin VB.Menu audioTrack 
         Caption         =   "音轨"
         Begin VB.Menu audioTrack0 
            Caption         =   "音轨 0"
         End
         Begin VB.Menu audioTrack1 
            Caption         =   "音轨 1"
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助"
      Begin VB.Menu homepage 
         Caption         =   "主页"
      End
      Begin VB.Menu pathCodec 
         Caption         =   "解码器路径"
      End
      Begin VB.Menu about 
         Caption         =   "关于"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Sub about_Click()
    MsgBox "海天鹰播放器 1.0" + vbCrLf + "基于 Aplayer 的媒体播放器" + vbCrLf + "http://aplayer.open.xunlei.com/", , "关于"
End Sub

Private Sub audioTrack0_Click()
    Player1.GetAPlayerObject.SetConfig 403, 0
    audioTrack0.Checked = True
    audioTrack1.Checked = False
End Sub

Private Sub audioTrack1_Click()
    Player1.GetAPlayerObject.SetConfig 403, 1
    audioTrack0.Checked = False
    audioTrack1.Checked = True
End Sub

Private Sub backward_Click()
    Dim position
    position = Player1.GetAPlayerObject.GetPosition - 5000
    If position > 0 Then Player1.GetAPlayerObject.SetPosition (position)
End Sub

Private Sub capture_Click()
    Dim path
    path = App.path + "\Capture"
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    Player1.GetAPlayerObject.SetConfig 703, Player1.GetAPlayerObject.GetVideoWidth
    Player1.GetAPlayerObject.SetConfig 704, Player1.GetAPlayerObject.GetVideoHeight
    Player1.GetAPlayerObject.SetConfig 702, App.path + "\Capture\" + Format(Now, "YYYYMMDDHHmmss") + ".jpg"
End Sub

Private Sub flipH_Click()
    If Player1.GetAPlayerObject.GetConfig(302) = 0 Then
        Player1.GetAPlayerObject.SetConfig 302, 1
    Else
        Player1.GetAPlayerObject.SetConfig 302, 0
    End If
End Sub

Private Sub flipV_Click()
    If Player1.GetAPlayerObject.GetConfig(303) = 0 Then
        Player1.GetAPlayerObject.SetConfig 303, 1
    Else
        Player1.GetAPlayerObject.SetConfig 303, 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode
    If (KeyCode = 13) Then Call fullscreen_Click
End Sub

Private Sub Form_Load()
    Form1.Width = 800 * Screen.TwipsPerPixelX
    Form1.Height = 600 * Screen.TwipsPerPixelY
    fillList ("tv.txt")
    Player1.GetAPlayerObject.SetConfig 1305, App.path + "\杨宗纬&张碧晨-凉凉.lrc"
    Player1.GetAPlayerObject.SetConfig 1308, App.path + "\zbc.jpg"
    Player1.GetAPlayerObject.SetConfig 1310, 2
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 1 To Data.Files.Count '逐个读取文件路径
        Debug.Print Data.Files(i)
    Next
    Player1.GetAPlayerObject.open (Data.Files(1))
    List1.Visible = False
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then '排除最小化
        Player1.Width = Form1.Width
        Player1.Height = Form1.Height - 800
        List1.Height = Form1.Height - 1600
        List1.Left = Form1.Width - List1.Width
    End If
End Sub

Private Sub forward_Click()
    Dim position
    position = Player1.GetAPlayerObject.GetPosition + 5000
    If position < Player1.GetAPlayerObject.GetDuration Then Player1.GetAPlayerObject.SetPosition (position)
End Sub

Private Sub fullscreen_Click()
    If (Not Player1.IsFullScreen) Then Player1.SetFullScreen (True)
End Sub

Private Sub GIF_Click()
    Dim path
    path = App.path + "\Capture"
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    Player1.GetAPlayerObject.SetConfig 707, "4"
    Player1.GetAPlayerObject.SetConfig 703, "320"
    Player1.GetAPlayerObject.SetConfig 704, "240"
    Player1.GetAPlayerObject.SetConfig 709, "length=6000;cutinterval=1000;playinterval=1000"
    Player1.GetAPlayerObject.SetConfig 702, App.path + "\Capture\" + Format(Now, "YYYYMMDDHHmmss") + ".gif"
End Sub

Private Sub homepage_Click()
    Shell "explorer.exe https://github.com/sonichy"
End Sub

Private Sub info_Click()
    Dim fileSize
    Debug.Print TypeName(Player1.GetAPlayerObject.GetConfig(5)) + Player1.GetAPlayerObject.GetConfig(5)
    fileSize = SB(Val(Player1.GetAPlayerObject.GetConfig(5)))
    MsgBox "文件路径：" + Player1.GetAPlayerObject.GetConfig(4) + vbCrLf + "文件大小：" + fileSize + vbCrLf + " 分辨率 ：" + Str(Player1.GetAPlayerObject.GetVideoWidth()) + " X" + Str(Player1.GetAPlayerObject.GetVideoHeight()) + vbCrLf + " 音  频 ：" + Player1.GetAPlayerObject.GetConfig(402), , "海天鹰播放器"
End Sub

Private Sub List1_Click()
    Dim c As String
    c = List1.List(List1.ListIndex)
    If InStr(c, ",") Then
        Dim item() As String
        item = Split(c, ",")
        Dim surl
        surl = item(1)
        Debug.Print surl
        Player1.GetAPlayerObject.open (surl)
    End If
End Sub

Private Sub live_Click()
    If List1.Visible = False Then
        List1.Visible = True
    Else
        List1.Visible = False
    End If
End Sub

Private Sub mute_Click()
    If Player1.GetAPlayerObject.GetConfig(12) = 0 Then
        Player1.GetAPlayerObject.SetConfig 12, 1
         Me.Caption = "海天鹰播放器 - 静音"
    Else
        Player1.GetAPlayerObject.SetConfig 12, 0
        Me.Caption = "海天鹰播放器 - 音量" + Str(Player1.GetAPlayerObject.GetVolume())
    End If
End Sub

Private Sub open_Click()
    Dim ofn As OPENFILENAME
    Dim rtn As String
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "所有文件"
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    'ofn.lpstrInitialDir = App.Path
    ofn.lpstrTitle = "打开文件"
    ofn.flags = 6148
    rtn = GetOpenFileName(ofn)
    If rtn >= 1 Then
        Player1.GetAPlayerObject.open (ofn.lpstrFile)
    End If
End Sub

Private Sub openURL_Click()
    Dim surl
    SendKeys "{home}"
    surl = InputBox("请输入网址：", "海天鹰播放器", Clipboard.GetText)
    If surl <> "" Then
        Player1.GetAPlayerObject.open (surl)
        List1.Visible = False
    End If
End Sub

Private Sub pathCapture_Click()
    Shell "explorer.exe " & App.path + "\Capture", vbNormalFocus
End Sub

Private Sub pathCodec_Click()
    Shell "explorer.exe " + Player1.GetAPlayerObject.GetConfig(2), vbNormalFocus
End Sub

Private Sub Player1_OnDownloadCodec(ByVal strCodecPath As String)
    MsgBox ("缺少编码：" + strCodecPath)
End Sub

Private Sub Player1_OnMessage(ByVal nMessage As Long, ByVal wParam As Long, ByVal lParam As Long)
    'Debug.Print nMessage
End Sub

Private Sub Player1_OnOpenSucceeded()
    Me.Caption = "海天鹰播放器 - " + Player1.GetAPlayerObject.GetConfig(4)
    audioTrack0_Click
    List1.Visible = False
End Sub

Private Sub Player1_OnVideoSizeChanged()
    If Player1.GetAPlayerObject.GetVideoWidth > 0 And Player1.GetAPlayerObject.GetVideoHeight > 0 Then
        Form1.Width = Player1.GetAPlayerObject.GetVideoWidth * Screen.TwipsPerPixelX
        Form1.Height = Player1.GetAPlayerObject.GetVideoHeight * Screen.TwipsPerPixelY + 1100
    End If
End Sub

Private Sub quit_Click()
    End
End Sub

Private Sub soundTrack0_Click()

End Sub

Private Sub rotate_Click()
    Debug.Print Player1.GetAPlayerObject.GetConfig(304)
    Player1.GetAPlayerObject.SetConfig 304, Player1.GetAPlayerObject.GetConfig(304) + 180
End Sub

Private Sub chooseSubtitle_Click()
    Dim ofn As OPENFILENAME
    Dim rtn As String
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "所有文件"
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    'ofn.lpstrInitialDir = App.Path
    ofn.lpstrTitle = "选择字幕"
    ofn.flags = 6148
    rtn = GetOpenFileName(ofn)
    If rtn >= 1 Then
        Player1.GetAPlayerObject.SetConfig 503, ofn.lpstrFile
        switchSubtitle.Checked = True
    End If
End Sub


Private Sub switchSubtitle_Click()
    If Player1.GetAPlayerObject.GetConfig(504) = 0 Then
        Player1.GetAPlayerObject.SetConfig 504, 1
        switchSubtitle.Checked = True
    Else
        Player1.GetAPlayerObject.SetConfig 504, 0
        switchSubtitle.Checked = False
    End If
End Sub

Private Sub volumeDown_Click()
    Player1.GetAPlayerObject.SetConfig 12, 0
    Player1.GetAPlayerObject.SetVolume (Player1.GetAPlayerObject.GetVolume() - 1)
    Me.Caption = "海天鹰播放器 - 音量" + Str(Player1.GetAPlayerObject.GetVolume())
    Player1.GetAPlayerObject.SetConfig 606, 10
    Player1.GetAPlayerObject.SetConfig 607, 10
    Player1.GetAPlayerObject.SetConfig 612, "音量" + Str(Player1.GetAPlayerObject.GetVolume())
    Debug.Print "bound " + Player1.GetAPlayerObject.GetConfig(603)
End Sub

Private Sub volumeUp_Click()
    Player1.GetAPlayerObject.SetConfig 12, 0
    Player1.GetAPlayerObject.SetVolume (Player1.GetAPlayerObject.GetVolume() + 1)
    Me.Caption = "海天鹰播放器 - 音量" + Str(Player1.GetAPlayerObject.GetVolume())
    Player1.GetAPlayerObject.SetConfig 606, 10
    Player1.GetAPlayerObject.SetConfig 607, 10
    Player1.GetAPlayerObject.SetConfig 612, "音量" + Str(Player1.GetAPlayerObject.GetVolume())
    'Player1.GetAPlayerObject.SetConfig 622, 1
    Debug.Print "Sprite " + Player1.GetAPlayerObject.GetConfig(2301)
End Sub

Function fillList(filename As String)
    Dim s As String
    Dim filepath
    filepath = App.path + "\" + filename
    'Debug.Print filepath
    Open filepath For Binary As #1
    s = Space(LOF(1))
    Get #1, 1, s
    Close #1
    'Debug.Print s
    Dim line() As String
    line = Split(s, vbCrLf)
    For i = 0 To UBound(line)
        'If InStr(line(i), ",") Then
            'Dim tv() As String
            'tv = Split(line(i), ",")
            'List1.AddItem (tv(0))
            List1.AddItem (line(i))
        'End If
    Next
End Function

Function SB(bytes As Double) As String
    If bytes > 1000000000 Then
        SB = Format(bytes / 1024 ^ 3, "#.### GB")
    ElseIf bytes > 1000000 Then
        SB = Format(bytes / 1024 ^ 2, "#.### MB")
    ElseIf bytes > 1000 Then
        SB = Format(bytes / 1024, "#.### KB")
    Else
        SB = Str(bytes) + " B"
    End If
End Function
