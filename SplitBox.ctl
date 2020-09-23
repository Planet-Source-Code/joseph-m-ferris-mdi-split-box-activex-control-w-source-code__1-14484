VERSION 5.00
Begin VB.UserControl SplitBox 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer timResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3660
      Top             =   2985
   End
   Begin VB.Timer timCheck 
      Interval        =   100
      Left            =   4155
      Top             =   2985
   End
   Begin VB.PictureBox picHandle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   780
      MousePointer    =   9  'Size W E
      ScaleHeight     =   165
      ScaleWidth      =   1590
      TabIndex        =   0
      Top             =   75
      Width           =   1590
   End
End
Attribute VB_Name = "SplitBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pMinWidth As Long
Private pMaxWidth As Long

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub timCheck_Timer()

'*' Check the current alignment state of the Control Instance.

Static LastState As Integer     '*' Static to check last value in this Sub

'*' Only evaluate if the alinment has changed.
'*'
If LastState = UserControl.Extender.Align Then
    Exit Sub
Else
    LastState = UserControl.Extender.Align
End If

'*' Display based upon the current alignment.
'*'
Select Case UserControl.Extender.Align

Case 0      '*' None

    SetUpNone
    
Case 1      '*' Top

    SetUpTop

Case 2      '*' Bottom

    SetUpBottom
    
Case 3      '*' Left

    SetUpLeft
    
Case 4      '*' Right

    SetUpRight

End Select

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_Resize()
    
    RaiseEvent Resize
    
timCheck.Enabled = UserControl.Ambient.UserMode

'*' Once again, resizing is dependant upon the alignment of the Control.
'*'
Select Case UserControl.Extender.Align

Case 0      '*' None

    SetUpNone
    
Case 1      '*' Top

    SetUpTop

Case 2      '*' Bottom

    SetUpBottom
    
Case 3      '*' Left

    SetUpLeft
    
Case 4      '*' Right

    SetUpRight

End Select

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
End Sub

Private Sub SetUpTop()

    '*' Arrange control if it is on the top of the MDI Form.
    '*'
    '*' Set the pointer to NS Resize
    '*'
    picHandle.MousePointer = 7
    
    '*' Move into bottom left area of control
    '*'
    picHandle.Left = 0
    picHandle.Top = UserControl.ScaleHeight - (150 * (UserControl.ScaleHeight / UserControl.Height))
    
    '*' Render the size of the hidden picture box that controls the resize.
    '*'
    picHandle.Height = 150 * (UserControl.ScaleHeight / UserControl.Height)
    picHandle.Width = UserControl.Width

End Sub

Private Sub SetUpBottom()

    '*' Arrange control if it is on the top of the MDI Form.
    '*'
    '*' Set the pointer to NS Resize
    '*'
    picHandle.MousePointer = 7
    
    '*' Move into top left area of control
    '*'
    picHandle.Left = 0
    picHandle.Top = 0
    
    '*' Render the size of the hidden picture box that controls the resize.
    '*'
    picHandle.Height = 150 * (UserControl.ScaleHeight / UserControl.Height)
    picHandle.Width = UserControl.Width
    
End Sub

Private Sub SetUpLeft()

    '*' Arrange control if it is on the top of the MDI Form.
    '*'
    '*' Set the pointer to EW Resize
    '*'
    picHandle.MousePointer = 9
    
    '*' Move into the top right area of control
    '*'
    picHandle.Left = UserControl.ScaleWidth - (150 * (UserControl.ScaleWidth / UserControl.Width))
    picHandle.Top = 0
    
    '*' Render the size of the hidden picture box that controls the resize.
    '*'
    picHandle.Width = 150 * (UserControl.ScaleWidth / UserControl.Width)
    picHandle.Height = UserControl.Height

End Sub

Private Sub SetUpRight()

    '*' Arrange control if it is on the top of the MDI Form.
    '*'
    '*' Set the pointer to EW Resize
    '*'
    picHandle.MousePointer = 9

    '*' Move into the top left area of the control.
    '*'
    picHandle.Left = 0
    picHandle.Top = 0
    
    '*' Render the size of the hidden picture box that controls the resize.
    '*'
    picHandle.Width = 150 * (UserControl.ScaleWidth / UserControl.Width)
    picHandle.Height = UserControl.Height

End Sub

Private Sub SetUpNone()

    '*' There is no allowance for no alignment, refer to the right alignment.
    '*'
    SetUpRight
    
End Sub

Private Sub picHandle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'*' The MouseDown event in the picHandle object is the trigger that will allow the tracking
'*' of the mouse position to begin.  All of the calculations are done in a timer to allow
'*' for the repetative and constant tracking of the mouse position and object sizes.
'*'
timResize.Enabled = True

End Sub

Private Sub picHandle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'*' The MouseUp event kills the trigger that was set on the MouseDown event.  This will mean
'*' that the user has release the mouse button and does not wish to resize the split plane
'*' any further.
'*'
timResize.Enabled = False

End Sub

Private Sub timResize_Timer()

Dim MinWidth As Long            '*' Minimum Width of the Split Plane
Dim MaxWidth As Long            '*' Maximum Width of the Split Plane
Dim intCalc As Long             '*' Calculated value of the width of the Split Plane
Dim CurrentX As Long            '*' Current XValue of the Mouse
Dim CurrentY As Long            '*' Current YValue of the Mouse

Static LastMousePosX As Long    '*' Static Variable to Track the Mouse's Last Known Position
Static LastMousePosY As Long

'*' Get the current value of the X and Y based upon the location of the mouse.
'*'
CurrentX = GetX
CurrentY = GetY

'*' The minimum and maximum width of the splitter plane can be set to either an absolute value
'*' or to an equation (Future).  Values are represented in Twips.
'*'
MinWidth = 150                                    '*' On some machines, smaller values cause jumpiness.
MaxWidth = (UserControl.Parent.Width / 2) - 150   '*' Limit the maximum to be one half of the form's size.
MinHeight = 150
MaxHeight = (UserControl.Parent.Height / 2) - 150

'*' Equation for Determining the width of the PictureBox (Solve for i)
'*'
'*' Note: This equation is for the right hand box alignment only.  Varies for other
'*'       alignments.
'*'
'*' Mx = Left of the MDI Form
'*' Px = Left of the Mouse Pointer
'*' S = Scale (Width/ScaleWidth)
'*' Mw = Width of the MDI Form
'*'
'*' i = Mx - (Px * S) + Mw
'*'
'*' Yields: Anticipated width of the Split Plane
'*'
Select Case UserControl.Extender.Align

    Case 1      '*' Top

        '*' Because of the titlebar and possible menus, we can not refer to the top of the
        '*' form, since it might be off by a few hundred twips.  We need to create a RECT and
        '*' store the value of a GetWindowRect API Call to return the top of the control.
        '*'
        Dim rctUser As RECT
        
        '*' Get the bounding box of the control
        '*'
        r = GetWindowRect(UserControl.hwnd, rctUser)
        
        '*' Calculate proposed height
        '*'
        intCalc = (CurrentY * (UserControl.Width / UserControl.ScaleWidth)) - (rctUser.Top * (UserControl.Width / UserControl.ScaleWidth))

        '*' Bounds Checking
        If intCalc <= MinHeight Then
            intCalc = MinHeight
        ElseIf intCalc >= MaxHeight Then
            intCalc = MaxHeight
        End If

        '*' Set the height of the Split Plane to be equal to that of the value for the equation.
        '*'
        UserControl.Height = intCalc

    Case 2      '*' Bottom

    
        '*' Calculate proposed height
        '*'
        intCalc = UserControl.Parent.Top - (CurrentY * (UserControl.Width / UserControl.ScaleWidth)) + UserControl.Parent.Height

        '*' Bounds Checking
        '*'
        If intCalc <= MinHeight Then
            intCalc = MinHeight
        ElseIf intCalc >= MaxHeight Then
            intCalc = MaxHeight
        End If

        '*' Set the height of the Split Plane to be equal to that
        UserControl.Height = intCalc

    Case 3      '*' Left

        '*' Calculate the proposed width
        '*'
        intCalc = (CurrentX * (UserControl.Width / UserControl.ScaleWidth)) - UserControl.Parent.Left

        '*' Bounds Checking
        If intCalc <= MinWidth Then
            intCalc = MinWidth
        ElseIf intCalc >= MaxWidth Then
            intCalc = MaxWidth
        End If

        '*' Set the width of the Split Plane to be equal to that
        UserControl.Width = intCalc

    Case 4      '*' Right
    
        '*' Calculate the proposed width
        '*'
        intCalc = UserControl.Parent.Left - (CurrentX * (UserControl.Width / UserControl.ScaleWidth)) + UserControl.Parent.Width

        '*' Bounds Checking
        '*'
        If intCalc <= MinWidth Then
            intCalc = MinWidth
        ElseIf intCalc >= MaxWidth Then
            intCalc = MaxWidth
        End If

        '*' Set the width of the Split Plane to be equal to that
        UserControl.Width = intCalc

End Select



End Sub

