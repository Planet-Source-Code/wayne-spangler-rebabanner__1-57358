Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public FileName As String           'Name of Banner File
Public BannerPic As String
Public BackgroundPic As String
Public PictureWidth As Long
Public PictureHeight As Long
Public PictureLeft As Long
Public PictureTop As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public LiteDiff As Single
Public DarkDiff As Single
Public BorderColor As Long
Public CustomX As Long
Public CustomY As Long
Public Dirty As Boolean
Public Ratio As Single

'******************************************************************************************
'This changes the size of the banner. The size is passed in a string called "size" and is in the
'form:  "123 x 456".
'******************************************************************************************
Public Sub SetSize(size As String)
    'x and y are the width and height of the banner. Z is an index into the string.
    Dim x As Integer, y As Integer, z As Integer
    'First find the "x" in the string
    z = InStr(size, "x")
    'Now get the value of x by geting the value of what is to the left of the "x"
    x = Val(Left$(size, z - 1))
    CustomX = x
    'And the value of y to the right of the "x"
    y = Val(Mid$(size, z + 1))
    CustomY = y
    'Change the size of the "Display" picture box. We use move to keep it from flickering.
    With frmMain
        .Display.Move 8, 8, x, y
        'then refresh the "Display" picture box. This updates the image that is in memory.
        'Set .Display.Picture = .Display.Image
        .Display.Refresh
    End With
End Sub

'************************************************************************
'Redraws the banner.
'************************************************************************
Public Sub Redraw()
    With frmMain
        .Display.Cls
        'check if the bacdground graphic option button is set
        'if it is then tile the background on the banner
        If .optGraphic Then Tile
        'if there is a picture then put it on the banner
        If .DisplayPic Then ShowPicture
        'set the font
        .Display.FontName = .cboFontName.List(.cboFontName.ListIndex)
        .Display.FontBold = .txt1Bold
        .Display.FontItalic = .txt1Italic
        .Display.FontUnderline = .txt1Underline
        .Display.FontSize = .txt1FontSize
        'If there is a shadow then put it on the banner
        If .Shadow Then PrintShadow
        'if there is a 3d text then put it on the banner
        If .OutlineText Then PrintOutline
        'now set the print position for the regular text and print it
        .Display.CurrentY = Val(.txtVert)
        .Display.CurrentX = Val(.txtHorz)
        .Display.Print .Text1
        ' is a border being used
        If .ShowBorder Then PrintBorder
    End With
End Sub

'************************************************************************
' Prints the text in black offset by the value in .txtShadow
'************************************************************************
Private Sub PrintBorder()
    With frmMain
        ' a border is used
        If .LineBorder Then
            ' It is a line
            PrintLineBorder
        Else
            'it is a shaded border
            PrintShadeBorder
        End If
    End With
End Sub

'************************************************************************
' Prints a shaded border around the banner
'************************************************************************
Private Sub PrintShadeBorder()
    Dim x As Long, y As Long, y1 As Long
    ' a shaded border is to be drawn around the banner
    ' first check if the border width is not 0
    If Val(frmMain.BorderWidth) > 0 Then
        With frmMain
        '================= shade top ==================================
            For y = 0 To .BorderWidth
                For x = y To .Display.ScaleWidth - y
                    SetColor x, y, LiteDiff
                Next x
            Next y
        '================= shade left ==================================
            For x = 0 To .BorderWidth
                For y = x To .Display.ScaleHeight - x
                    SetColor x, y, LiteDiff
                Next y
            Next x
            '================ Shade Bottom =======================
            y1 = .BorderWidth
            For y = .Display.ScaleHeight - .BorderWidth - 1 To .Display.ScaleHeight
                For x = y1 To .Display.ScaleWidth - y1
                    SetColor x, y, DarkDiff
                Next x
                y1 = y1 - 1
            Next y
        '================= shade right ==================================
            y1 = 0
            For x = .Display.ScaleWidth To .Display.ScaleWidth - .BorderWidth - 1 Step -1
                For y = y1 To .Display.ScaleHeight - y1
                    SetColor x, y, DarkDiff
                Next y
                y1 = y1 + 1
            Next x
        End With
    End If
End Sub

'************************************************************************
' Prints a solid border around the banner
'************************************************************************
Private Sub SetColor(x As Long, y As Long, Diff As Single)
    Dim color As Long, Red As Integer, Green As Integer, Blue As Integer
    With frmMain
        ' Break a color into its components.
         color = GetPixel(.Display.hdc, x, y)
         If color <> &HFFFFFFFF Then
             If color And &H80000000 Then color = GetSysColor(color And &HFFFFFF)
             Red = color And &HFF&
             Green = (color And &HFF00&) \ &H100&
             Blue = (color And &HFF0000) \ &H10000
         End If
         ' adjust the colors
         Red = Red * Diff
         Green = Green * Diff
         Blue = Blue * Diff
         ' and check it the colors are right
         CheckColors Red, Green, Blue
         SetPixel .Display.hdc, x, y, RGB(Red, Green, Blue)
    End With
End Sub

'************************************************************************
' Prints a solid border around the banner
'************************************************************************
Private Sub PrintLineBorder()
    Dim x As Long, y As Long
    With frmMain
        'a line is to be drawn around the banner
        ' and will be borderwidth wide
        ' start with top
        For y = 0 To Val(.BorderWidth)
            .Display.Line (0, y)-(.Display.ScaleWidth, y), BorderColor
        Next y
        'now the bottom
        For y = .Display.ScaleHeight - Val(.BorderWidth) - 1 To .Display.ScaleHeight
            .Display.Line (0, y)-(.Display.ScaleWidth, y), BorderColor
        Next
        'now the left
        For x = 0 To Val(.BorderWidth)
            .Display.Line (x, 0)-(x, .Display.ScaleHeight), BorderColor
        Next
        'and then the right
        For x = .Display.ScaleWidth - Val(.BorderWidth) - 1 To .Display.ScaleWidth
            .Display.Line (x, 0)-(x, .Display.ScaleHeight), BorderColor
        Next
    End With
End Sub

'************************************************************************
' Prints the text in black offset by the value in .txtShadow
'************************************************************************
Private Sub PrintShadow()
    Dim HoldColor As Long
    With frmMain
        'to get the black text to print we have to change the display forecolor
        'to black but we want to matain the forecolor so save it
        HoldColor = .Display.ForeColor
        .Display.ForeColor = vbBlack
        ' set the print position using the offset
        .Display.CurrentY = Val(.txtVert) + Val(.txtShadow)
        .Display.CurrentX = Val(.txtHorz) + Val(.txtShadow)
        ' and print the text
        .Display.Print .Text1
        'recover the saved forecolor
        .Display.ForeColor = HoldColor
    End With
End Sub

'************************************************************************
' Prints the text in black offset by the value in .txtShadow
'************************************************************************
Private Sub PrintOutline()
    Dim HoldColor As Long
    With frmMain
        'to get the black text to print we have to change the display forecolor
        'to black but we want to matain the forecolor so save it
        HoldColor = .Display.ForeColor
        .Display.ForeColor = .OutlineColor.BackColor
        ' set the print position up one left one
        .Display.CurrentY = Val(.txtVert) - 2
        .Display.CurrentX = Val(.txtHorz) - 2
        ' and print the text
        .Display.Print .Text1
         ' set the print position down one left one
        .Display.CurrentY = Val(.txtVert) + 2
        .Display.CurrentX = Val(.txtHorz) - 2
        ' and print the text
        .Display.Print .Text1
        ' set the print position up one right one
        .Display.CurrentY = Val(.txtVert) - 2
        .Display.CurrentX = Val(.txtHorz) + 2
        ' and print the text
        .Display.Print .Text1
        ' set the print position down one right one
        .Display.CurrentY = Val(.txtVert) + 2
        .Display.CurrentX = Val(.txtHorz) + 2
        ' and print the text
        .Display.Print .Text1
       'recover the saved forecolor
        .Display.ForeColor = HoldColor
    End With
End Sub

'************************************************************************
' Transfers the picture from the hold control to the banner
' The aspect ration is kept
'************************************************************************
Private Sub ShowPicture()
    With frmMain
        If .DisplayPic Then
            If .BanPic.Picture.Type = vbPicTypeNone Then Exit Sub
            'maintain the aspect raion
            PictureWidth = .BanPic.Width * PictureHeight / .BanPic.Height
            'Stretch image to banner
            .Display.PaintPicture .BanPic.Picture, _
                                    PictureLeft, PictureTop, _
                                    PictureWidth, PictureHeight, _
                                    0, 0, _
                                    .BanPic.Width, .BanPic.Height
        End If
    End With
End Sub

'************************************************************************
'Tiles the background image on the "Display"
'************************************************************************
Public Sub Tile()
    'x is the horzontal position to place the graphic
    'y is the vertical position to place the graphic
    Dim x As Long, y As Long
    'If there is no picture in the graphic an error will accure so we
    'need to check. If no picture then exit
    If frmMain.PicBG.Picture.Type = vbPicTypeNone Then Exit Sub
    With frmMain
        'The x position is moved the width of the picture each time through.
        For x = 0 To .Display.ScaleWidth Step .PicBG.ScaleWidth
            'The y position is moved the height of the picture each time throught
            For y = 0 To .Display.ScaleHeight Step .PicBG.ScaleHeight
                'Now paint the graphic on to the display at the x,y coordenance
                .Display.PaintPicture .PicBG.Picture, x, y
            Next y
        Next x
    End With
End Sub

'************************************************************************
' Checks to see if any of the colors are out of range and corrects them
'************************************************************************
Private Sub CheckColors(r As Integer, g As Integer, b As Integer)
    If r > 255 Then r = 255
    If r < 0 Then r = 0
    If g > 255 Then g = 255
    If g < 0 Then g = 0
    If b > 255 Then b = 255
    If b < 0 Then b = 0
End Sub
