VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam As String) As Long

Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const LB_DELETESTRING = &H182
Private Const LB_GETCURSEL = &H188
Private Const LB_INSERTSTRING = &H181
Private Const LB_GETCOUNT = &H18B
Private Const LB_GETTEXT = &H189
Private Const LB_SETCURSEL = &H186
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const LB_ITEMFROMPOINT = &H1A9

Private LB As ListBox
Private LBBound As Boolean
Private m_Scroll As Boolean

Public Function BindList(LBox As ListBox) As Boolean
    Set LB = LBox
    If LB.hWnd > 0 Then
        BindList = True: LBBound = True
    Else
        BindList = False: LBBound = False
    End If
End Function

Public Property Get Text() As String
    Dim sItemText As String * 255
    Dim FinalOutStr As String
    
    sItemText = ""
    Call SendMessageStr(LB.hWnd, LB_GETTEXT, Me.ListIndex, ByVal sItemText)
    FinalOutStr = RTrim(sItemText)
    If Len(FinalOutStr) > 0 Then FinalOutStr = Left(FinalOutStr, Len(FinalOutStr) - 1)
    Text = FinalOutStr
End Property

Public Property Get List(Index As Long) As String
    Dim sItemText As String * 255
    Dim FinalOutStr As String
    
    sItemText = ""
    Call SendMessageStr(LB.hWnd, LB_GETTEXT, Index, ByVal sItemText)
    FinalOutStr = RTrim(sItemText)
    If Len(FinalOutStr) > 0 Then FinalOutStr = Left(FinalOutStr, Len(FinalOutStr) - 1)
    List = FinalOutStr
End Property

Public Property Get ListCount() As Long
    ListCount = SendMessageLong(LB.hWnd, LB_GETCOUNT, 0&, 0&)
End Property

Public Property Get ListIndex() As Long
    ListIndex = SendMessage(LB.hWnd, LB_GETCURSEL, 0, 0)
End Property

Public Property Let ListIndex(Index As Long)
    Call SendMessage(LB.hWnd, LB_SETCURSEL, Index, 0)
End Property

Public Sub AddItem(Item As String, Optional Index As Long = -1, Optional CheckForDuplicateEntry As Boolean = False)
    Dim DupEntry As Long
    If Index = -1 Then Index = Me.ListCount
    If CheckForDuplicateEntry Then
        DupEntry = SendMessageByString&(LB.hWnd, LB_FINDSTRINGEXACT, 0, Item)
        If DupEntry = -1 Then Exit Sub
    End If
    Call SendMessage(LB.hWnd, LB_INSERTSTRING, Index, ByVal Item)
End Sub

Public Sub ModifyItem(Index As Long, Item As String)
    Me.RemoveItem Index
    Call SendMessage(LB.hWnd, LB_INSERTSTRING, Index, ByVal Item)
End Sub

Public Sub RemoveItem(Index As Long)
    Call SendMessage(LB.hWnd, LB_DELETESTRING, Index, 0)
End Sub

Public Function LoadListFromFile(FilePath As String, Optional AllowDups As Boolean = True)
    Dim sLineIn As String
    Dim DupEntry As Long
    
    On Error GoTo ErrLoadListFromFile
    
    Open FilePath For Input As #1
    While Not EOF(1)
        If AllowDups = True Then
            Line Input #1, sLineIn
            If Trim$(sLineIn) <> "" Then Me.AddItem sLineIn
        Else
            Line Input #1, sLineIn
            DupEntry = SendMessageByString&(LB.hWnd, LB_FINDSTRINGEXACT, 0, Trim$(sLineIn))
            If Trim$(sLineIn) <> "" And DupEntry = -1 Then Me.AddItem sLineIn
        End If
    Wend
    Close #1
    FilePath = ""
AfterLoadListFromFile:
Exit Function

ErrLoadListFromFile:
    Resume AfterLoadListFromFile
End Function

Public Sub SaveListToFile(OutputFile As String, Optional AppendFile As Boolean = False)
    On Error GoTo Error_Killer
    Dim Nbr As Long
    Dim sItemText As String * 255
    Dim FinalOutStr As String
    
    If AppendFile Then
        Open OutputFile For Append As #1
    Else
        Open OutputFile For Output As #1
    End If
    
    For Nbr = 0 To Me.ListCount - 1
        sItemText = ""
        Call SendMessageStr(LB.hWnd, LB_GETTEXT, Nbr, ByVal sItemText)
        FinalOutStr = RTrim(sItemText)
        FinalOutStr = Left(FinalOutStr, Len(FinalOutStr) - 1)
        Print #1, FinalOutStr
        Debug.Print Len(FinalOutStr)
    Next Nbr
    
    Close #1
    
    Exit Sub
Error_Killer:
    Close #1
    Exit Sub
End Sub

Private Sub UnlockControls()
    Call LockWindowUpdate(0&)
End Sub

Private Sub LockControls()
    Call LockWindowUpdate(LB.hWnd)
End Sub

Public Sub ExtendToolTip(Button As Integer, X As Single, Y As Single)
    Dim ListXPos  As Long
    Dim ListYPos  As Long
    Dim LBoxIndex As Long
    Dim sItemText As String * 255
    Dim FinalOutStr As String
    
    If Button = 0 Then
        ListXPos = CLng(X / Screen.TwipsPerPixelX)
        ListYPos = CLng(Y / Screen.TwipsPerPixelY)
    
        LBoxIndex = SendMessage(LB.hWnd, LB_ITEMFROMPOINT, 0, ByVal ((ListYPos * 65536) + ListXPos))
    
        If (LBoxIndex >= 0) And (LBoxIndex <= Me.ListCount - 1) Then
            sItemText = ""
            Call SendMessageStr(LB.hWnd, LB_GETTEXT, LBoxIndex, ByVal sItemText)
            FinalOutStr = RTrim(sItemText)
            If Len(FinalOutStr) > 1 Then FinalOutStr = Left(FinalOutStr, Len(FinalOutStr) - 1)
            LB.ToolTipText = FinalOutStr
        Else
            LB.ToolTipText = ""
        End If
    End If
End Sub

Public Sub RemoveDuplicateEntries(Optional FastLag As Boolean = False)
    On Error Resume Next
    Dim CurCount As Long
    Dim CurPos As Long
    Dim DupEntry As Long
    Dim sItemText As String * 255
    Dim FinalOutStr As String
    
    CurCount = Me.ListCount
    If CurCount = 0 Then Exit Sub
    CurPos = 0
    While CurPos < CurCount
        sItemText = "": FinalOutStr = ""
        Call SendMessageStr(LB.hWnd, LB_GETTEXT, CurPos, ByVal sItemText)
        FinalOutStr = RTrim(sItemText)
        If Len(FinalOutStr) > 0 Then FinalOutStr = Left(FinalOutStr, Len(FinalOutStr) - 1)
        If FinalOutStr <> "" Then
            DupEntry = SendMessageByString&(LB.hWnd, LB_FINDSTRINGEXACT, 0, FinalOutStr)
            If DupEntry <> CurPos Then
                Call SendMessage(LB.hWnd, LB_DELETESTRING, CurPos, 0)
                CurCount = CurCount - 1
            Else
                CurPos = CurPos + 1
            End If
        Else
            CurPos = CurPos + 1
        End If
        If FastLag = False Then DoEvents
    Wend
End Sub

Public Sub RightClickMenu(MForm As Form, PopMenu As Menu, X As Single, Y As Single)
    Dim lRet As Long
    Dim lXPos As Long, lYPos As Long

    lXPos = CLng(X / Screen.TwipsPerPixelX)
    lYPos = CLng(Y / Screen.TwipsPerPixelY)

    lRet = SendMessage(LB.hWnd, LB_ITEMFROMPOINT, 0, ByVal _
    ((lYPos * 65536) + lXPos))

    If lRet < Me.ListCount Then
        Me.ListIndex = lRet
        MForm.PopupMenu PopMenu
    End If
End Sub

Public Property Get HorizontalScroll() As Boolean
    HorizontalScroll = m_Scroll
End Property

Public Property Let HorizontalScroll(m_New_HScroll As Boolean)
    Dim i As Long, intGreatestLen As Long, lngGreatestWidth As Long
    If m_New_HScroll Then
        m_Scroll = True
        For i = 0 To Me.ListCount - 1
            Debug.Print Len(Me.List(i))
            If Len(Me.List(i)) > Len(Me.List(intGreatestLen)) Then
                intGreatestLen = i
            End If
        Next i
        lngGreatestWidth = LB.Parent.TextWidth(Me.List(intGreatestLen) + Space(1))
        lngGreatestWidth = LB.Parent.ScaleX(lngGreatestWidth, LB.Parent.ScaleMode, vbPixels)
        SendMessage LB.hWnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
    Else
        m_Scroll = False
        SendMessage LB.hWnd, LB_SETHORIZONTALEXTENT, 0, 0
    End If
End Property

Public Function Clear()
    LB.Clear
End Function
