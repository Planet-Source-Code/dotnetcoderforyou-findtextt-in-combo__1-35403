Attribute VB_Name = "Module1"
Declare Function SendMessage Lib "user32" Alias _
                                 "SendMessageA" _
                                 (ByVal hWnd As Long, _
                                  ByVal wMsg As Long, _
                                  ByVal wParam As Long, _
                                  lParam As Any) As Long
 Const CB_FINDSTRING = &H14C
 Const CB_ERR = (-1)


