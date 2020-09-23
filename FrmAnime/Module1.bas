Attribute VB_Name = "mdlAnime"
'============================
'By     Jim Jose
'email  jimjosev33@yahoo.com
'============================

'PLEASE READ THIS

'If you ( Feel Satisfactory )
'   Please 'Rate' this code.
'Else
'   Give feedback to improve this code.
'End If
'Good luck
'============================

Option Explicit

Private Type RECT   'Rectangle coordinates
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum AnimeEvent  'Determines the Animation on Loading/Unloading
    aUnload = 0
    aLoad = 1
End Enum

Public Enum AnimeSpeed  'Determines the Speed of animation
    aFast = 1
    aMedium = 10
    aSlow = 30
End Enum

Public Enum AnimeType   'Determines the choosed animation style
    aCentre = 0         'I think there will be more effects that you could add
    aLeftTop = 1
    aRightTop = 2
    aLeftBottom = 3
    aRightBottom = 4
End Enum

Private Type ICONINFO   'Getting the 'IconInfo'  of the Animating form
    fIcon As Long   'These values are needed to draw the icon in correct position
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

'Controll/Info API's Used
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long  'Gets the hdc of Desktop
Private Declare Function GetDesktopWindow Lib "user32" () As Long   'Gets the hwnd of Desktop
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long  'Gets Icon 'HotSpots'
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long) 'Controlls the 'Speed' of 'Loop'
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Draw/Clear API's Used
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal hIcon As Long) As Long   'Draw the icon on Screen
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long     'Clear up the screen
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long  'Draws Rectagle
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long     'Draw Animated rectangles( Using as the last event of animation )
'----------------------------------------------------------------------------------------------------------------------
'< Ref >
    'This Module mainly uses 'DrawAnimatedRects' API
    'This was my primary project(API) and all other are later Addings
    'The new feature is that the form 'Orginates' and 'Terminates' in it's own icon.( exicute to get this )
'< Info >
    '1)There is more possible 'Styles' that you can add.
    'Now the function contains only the standard moves
    '2)I don't know if sub 'TransRectangle" is a standard way.
    'I  have to do that since there is no direct linedrawing API/ The hollow rectangle drawing (by API) is more complex
'< Tips >
    '1)Change the 'Tolarence' value (in the sub 'AnimateForm') to set the icon in position  you want
    'Increasing the  'Tolarence' ,the icon will move to centre.
    '2)Change the 'DrawWidth' for the sub 'TransRectangle' for some different effects
    'now it is set to the default 'One'
    '3)Change the default 'RctCount' (Rectangle count) in the sub  'PrivateAnime'
'----------------------------------------------------------------------------------------------------------------------

'Animtion using 'DrawAnimatedRects' API
Public Sub AnimateForm(Frm As Form, aEvent As AnimeEvent, Optional aType As AnimeType = 0, _
                        Optional aSpeed As AnimeSpeed = 10, Optional SleepTime As Integer = 1)
Dim ScrX        As Long    'Determines the 'Screen.TwipsPerPixelX'
Dim ScrY        As Long    'Determines the 'Screen.TwipsPerPixelY'
Dim Icn         As ICONINFO 'Gets icon 'HotSpot'
Dim Rct1        As RECT    'The starting rect in 'Load' event
Dim Rct2        As RECT    'The ending rect in 'Load' event
Dim Tolarence   As Integer    'Determines the position of icon on the screen
Dim IconPosX    As Long    'Determines the position to icon to draw
Dim IconPosY    As Long

Tolarence = 50      'Increasing the value will move the icon to centre
ScrX = Screen.TwipsPerPixelX    'Setting value
ScrY = Screen.TwipsPerPixelY    'Setting value
GetIconInfo Frm.Icon, Icn       'Getting the 'IconInfo' to the variable 'Icn'

With Rct1   'Setting the First(Starting) rectangle as the dimensions of the form
    .Left = Frm.Left / ScrX     'Setting value
    .Top = Frm.Top / ScrY       'Setting value
    .Right = (Frm.Left + Frm.Width) / ScrX  'Setting value
    .Bottom = (Frm.Top + Frm.Height) / ScrY 'Setting value
End With

Select Case aType   'Selecting the case 'AnimationType'
    Case 0  '( Centre )
        With Rct2   'Setting values to the centre of the form
            .Left = (Rct1.Right + Rct1.Left) / 2 - Tolarence
            .Top = (Rct1.Bottom + Rct1.Top) / 2 - Tolarence
            .Right = Rct2.Left
            .Bottom = Rct2.Top
        End With
    Case 1  '( LeftTop )
        With Rct2   'Setting values to the 'LeftTop'
            .Left = Tolarence
            .Top = Tolarence
            .Right = Tolarence
            .Bottom = Tolarence
        End With
    Case 2  'RightTop
        With Rct2   'Setting values to the 'RightTop'
            .Left = Screen.Width / ScrX - Tolarence
            .Top = Tolarence
            .Right = Screen.Width / ScrX - Tolarence
            .Bottom = Tolarence
        End With
    Case 3  'LeftBottom
        With Rct2   'Setting values to the 'LeftBottom'
            .Left = Tolarence
            .Top = Screen.Height / ScrY - Tolarence
            .Right = Tolarence
            .Bottom = Screen.Height / ScrY - Tolarence
        End With
    Case 4  'RightBottom
        With Rct2   'Setting values to the 'RightBottom'
            .Left = Screen.Width / ScrX - Tolarence
            .Top = Screen.Height / ScrY - Tolarence
            .Right = Screen.Width / ScrX - Tolarence
            .Bottom = Screen.Height / ScrY - Tolarence
        End With
    'You can add more effects here
    End Select

IconPosX = (Rct2.Left + Rct2.Right) / 2 - Icn.xHotspot / ScrX 'Setting Icon X pos
IconPosY = (Rct2.Top + Rct2.Bottom) / 2 - Icn.yHotspot / ScrY 'Setting Icon Y pos
DrawIcon DeskDc, IconPosX, IconPosY, Frm.Icon   'Drawing the icon

If aEvent = 1 Then  'Load
    PrivateAnime Rct2, Rct1, aSpeed     'The Animation coded by me ( not API animation ) to draw with hollow rectangles
    DrawAnimatedRects Frm.hwnd, 3, Rct2, Rct1   'The API animation
End If

If aEvent = 0 Then  'Unload
    PrivateAnime Rct1, Rct2, aSpeed     'The Animation coded by me ( not API animation ) to draw with hollow rectangles
    DrawAnimatedRects Frm.hwnd, 3, Rct1, Rct2   'The API animation
    'The bellow code is used to set the icon there after unloading  for 1 sec
    Frm.Visible = False
    Unload Frm  'Unloading the form in the case of 'Unload' event
    ClearScreen
    DoEvents
    DrawIcon DeskDc, IconPosX, IconPosY, Frm.Icon   'ReDrawing the icon
    Sleep 1000 * SleepTime 'Sleeping  for 1 sec
End If
ClearScreen 'Clearing the Screen before exiting
DeleteObject Icn.fIcon
End Sub

'Returns the Desktop HDC
Private Function DeskDc()
    DeskDc = GetWindowDC(GetDesktopWindow)
End Function

'Returns the DeskTop Hwnd
Private Function DeskHwnd()
    DeskHwnd = GetDesktopWindow
End Function

'Clearing the sceen
Public Sub ClearScreen()
   InvalidateRect 0&, 0&, True
End Sub

'My Animation
Public Function PrivateAnime(sRct As RECT, eRct As RECT, ByVal aSpeed As AnimeSpeed, Optional ByVal RctCount = 25)
Dim X As Integer
Dim XIncr As Double
Dim YIncr As Double
Dim HIncr As Double
Dim WIncr As Double
Dim TempRect As RECT    'Declaring a 'Temporary rectagle' the dimensions in b/w the starting and ending rectangles

    XIncr = (eRct.Left - sRct.Left) / RctCount    'Determines Amount of change in each loop for the 'Left' property
    YIncr = (eRct.Top - sRct.Top) / RctCount    'Determines Amount of change in each loop for the 'Top' property
    HIncr = ((eRct.Bottom - eRct.Top) - (sRct.Bottom - sRct.Top)) / RctCount   'Determines Amount of change in each loop for the 'Height' of rectagle
    WIncr = ((eRct.Right - eRct.Left) - (sRct.Right - sRct.Left)) / RctCount    'Determines Amount of change in each loop for the 'Width' of rectagle
    TempRect = sRct
    
    For X = 1 To RctCount 'Doing the animation
        Sleep aSpeed    'Controlling the speed
        'Setting the Temporary rectangle's dimensions
        TempRect.Left = TempRect.Left + XIncr: TempRect.Right = TempRect.Right + XIncr + WIncr
        TempRect.Top = TempRect.Top + YIncr: TempRect.Bottom = TempRect.Bottom + YIncr + HIncr
        TransRectangle DeskDc, TempRect 'Drawing the Hollow rectangle
    Next X
End Function

'My Hollow rectangle drawing method ( I don't know if there is a standard method(API) )
'I have to do this because there was no direct line drawing API ,I could find.

'This sub created four other rectangles as the sides of the 'Required Rectangle'
'drawing all the four rectangle will result in the 'Required Rectangle'
Public Sub TransRectangle(Dhdc As Long, VRct As RECT, Optional ByVal DrawWidth As Long = 1)
Dim X As Integer
Dim TempRect(1 To 4) As RECT
    For X = 1 To 4
        TempRect(X) = VRct
        If X = 1 Then TempRect(X).Bottom = TempRect(X).Top + DrawWidth
        If X = 2 Then TempRect(X).Left = TempRect(X).Right - DrawWidth
        If X = 3 Then TempRect(X).Top = TempRect(X).Bottom - DrawWidth
        If X = 4 Then TempRect(X).Right = TempRect(X).Left + DrawWidth
        Rectangle Dhdc, TempRect(X).Left, TempRect(X).Top, TempRect(X).Right, TempRect(X).Bottom    'drawing the required rectangle
    Next X
End Sub
