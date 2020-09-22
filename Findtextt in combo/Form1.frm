VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   432
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Authors"
      Top             =   1920
      Width           =   1332
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Author"
      DataSource      =   "Data1"
      Height          =   288
      ItemData        =   "Form1.frx":0000
      Left            =   360
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   2892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function AutoFind(ByRef cboCurrent As ComboBox, _
                         ByVal KeyAscii As Integer, _
                         Optional ByVal LimitToList As Boolean = False)
        
Dim lCB As Long
Dim sFindString As String

'On Error GoTo Err_Handler
    If KeyAscii = 8 Then
        If cboCurrent.SelStart <= 1 Then
            cboCurrent = ""
            AutoFind = 0
            Exit Function
        End If
        If cboCurrent.SelLength = 0 Then
            sFindString = UCase(Left(cboCurrent, Len(cboCurrent) - 1))
        Else
            sFindString = Left$(cboCurrent.Text, cboCurrent.SelStart - 1)
        End If
    ElseIf KeyAscii < 32 Or KeyAscii > 127 Then
        Exit Function
    Else
        If cboCurrent.SelLength = 0 Then
            sFindString = UCase(cboCurrent.Text & Chr$(KeyAscii))
        Else
            sFindString = Left$(cboCurrent.Text, cboCurrent.SelStart) & Chr$(KeyAscii)
        End If
    End If
    lCB = SendMessage(cboCurrent.hWnd, CB_FINDSTRING, -1, ByVal sFindString)

    If lCB <> CB_ERR Then
        cboCurrent.ListIndex = lCB
        cboCurrent.SelStart = Len(sFindString)
        cboCurrent.SelLength = Len(cboCurrent.Text) - cboCurrent.SelStart
        AutoFind = 0
    Else
        If LimitToList = True Then
            AutoFind = 0
        Else
            AutoFind = KeyAscii
        End If
    End If
        
End Function
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoFind(Combo1, KeyAscii, False)

End Sub

Private Sub Form_Load()
Set dbNew1 = OpenDatabase(App.Path & "\Biblio.mdb")
Set Loay = dbNew1.OpenRecordset("Authors")
Do Until Loay.EOF
   Combo1.AddItem Loay!Author
   Loay.MoveNext
   Loop

End Sub
