Attribute VB_Name = "modFunctions"
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


'*' The POINTAPI Type contains two properties.  The first is the X Property, which returns
'*' the Horizontal Postion of the mouse as a long value, and the Y Property, which returns
'*' the Vertical Postion of the mouse as a long value.
'*'
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Function GetX() As Long

Dim POS As POINTAPI

'*' Call the GetCursorPos Function with the POINTAPI object that has been created.
'*'
GetCursorPos POS

'*' The value of X is now assigned to the X Property of the POS Object.  Assign this value
'*' to the function and return the value.
'*'
GetX = POS.x

End Function

Public Function GetY() As Long

Dim POS As POINTAPI

'*' Call the GetCursorPos Function with the POINTAPI object that has been created.
'*'
GetCursorPos POS

'*' The value of X is now assigned to the X Property of the POS Object.  Assign this value
'*' to the function and return the value.
'*'
GetY = POS.y

End Function


