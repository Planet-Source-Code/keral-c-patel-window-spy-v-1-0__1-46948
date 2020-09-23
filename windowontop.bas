Attribute VB_Name = "stayontop"
'This is/ Programmed by Keral. C. Patel./
'Date:- 8/1/2003
Private Declare Function SetWindowPos Lib "user32" _
         (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1


Public Function PutWindowOnTop(pFrm As Form)
  Dim lngWindowPosition As Long
  
  lngWindowPosition = SetWindowPos(pFrm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Function
