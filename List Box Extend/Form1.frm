VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   540
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Clear List"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   2940
      Width           =   1755
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Load List"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   2340
      Width           =   1755
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save List"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   2640
      Width           =   1755
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Remove Selected"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2040
      Width           =   1755
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Modify Item"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   900
      Width           =   1755
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove All Dups"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   1740
      Width           =   1755
   End
   Begin VB.CheckBox ckDup 
      Caption         =   "Check For Dup"
      Height          =   195
      Left            =   2880
      TabIndex        =   7
      Top             =   1500
      Width           =   1755
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Insert Item"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Item"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   600
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Index"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   300
      Width           =   1755
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   60
      Top             =   60
   End
   Begin VB.CheckBox ckHScroll 
      Caption         =   "Horizontal Scroll"
      Height          =   195
      Left            =   2880
      TabIndex        =   1
      Top             =   60
      Width           =   1755
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "Form1.frx":0000
      Left            =   60
      List            =   "Form1.frx":0028
      TabIndex        =   0
      Top             =   60
      Width           =   2775
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   2040
      Width           =   2820
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSIndex 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   1800
      Width           =   2820
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSelected 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuM1 
         Caption         =   "Menu Item 1"
      End
      Begin VB.Menu mnuM2 
         Caption         =   "Menu Item 2"
      End
      Begin VB.Menu mnuM3 
         Caption         =   "Menu Item 3"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cLB As New clsLB


Private Sub ckHScroll_Click()
    If ckHScroll.Value = Checked Then
        cLB.HorizontalScroll = True
    Else
        cLB.HorizontalScroll = False
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim Answere As Long
    Answere = InputBox("Enter New Index:", "Set Index", 0)
    cLB.ListIndex = Answere
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim Answere As Long
    Answere = InputBox("Enter Index To Find:", "Find Item", 0)
    MsgBox "The Item with the index of " & Answere & " is:" & vbCrLf & vbCrLf & cLB.List(Answere)
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim Answere(1) As String
    Answere(0) = InputBox("Enter Item Text:", "Insert Item")
    Answere(1) = InputBox("Enter Item Index (Optional):", "Insert Item")
    If Answere(1) = "" Then
        cLB.AddItem Answere(0), CLng(Answere(1)), ckDup.Value
    Else
        cLB.AddItem Answere(0), , ckDup.Value
    End If
End Sub

Private Sub Command4_Click()
    cLB.RemoveDuplicateEntries
End Sub

Private Sub Command5_Click()
    If cLB.ListIndex = -1 Then
        MsgBox "You must select an item to modify.", vbInformation, "Error"
        Exit Sub
    End If
    On Error Resume Next
    Dim Answere As String
    Answere = InputBox("Enter New Text:", "Modify Item")
    cLB.ModifyItem cLB.ListIndex, Answere
End Sub

Private Sub Command6_Click()
    If cLB.ListIndex = -1 Then
        MsgBox "You must select an item to remove.", vbInformation, "Error"
        Exit Sub
    End If
    cLB.RemoveItem cLB.ListIndex
End Sub

Private Sub Command7_Click()
    On Error GoTo ErrHand
    With CD1
        .DialogTitle = "Save List"
        .Filter = "Text Files|*.txt"
        .CancelError = True
        .ShowSave
        If Dir(.FileName) <> "" Then
            ans = MsgBox("The file " & .FileTitle & " already exists." & vbCrLf & _
            "Are you sure you want to overwrite it?", vbYesNo, "Confirm Overwrite")
            If ans = vbNo Then Exit Sub
        End If
        cLB.SaveListToFile .FileName
    End With
ErrHand:
End Sub

Private Sub Command8_Click()
    On Error GoTo ErrHand
    With CD1
        .DialogTitle = "Load List"
        .Filter = "Text Files|*.txt"
        .CancelError = True
        .ShowOpen
        cLB.LoadListFromFile .FileName
    End With
ErrHand:
End Sub

Private Sub Command9_Click()
    cLB.Clear
End Sub

Private Sub Form_Load()
    'Bind List1 to the class module
    'so that it knows wich listbox to use
    'For multiple listbox's you will need
    'to Dimension multiple classes
    'Dim cLB2 As New clsLB, cLB3 as New clsLB, etc.
    cLB.BindList List1
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cLB.ExtendToolTip Button, X, Y
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then cLB.RightClickMenu Me, mnuMenu, X, Y
End Sub

Private Sub Timer1_Timer()
    lblSelected = "Selected: " & cLB.Text
    lblSIndex = "Index: " & cLB.ListIndex
    lblCount = "Count: " & cLB.ListCount
    lblSIndex.Top = lblSelected.Top + lblSelected.Height
    lblCount.Top = lblSIndex.Top + lblSIndex.Height
End Sub
