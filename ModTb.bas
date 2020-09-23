Attribute VB_Name = "ModTb"
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwnewlong As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long

Private Const GCL_HBRBACKGROUND As Long = -10

Private Function GDI_TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, GDI_TranslateColor) Then
        GDI_TranslateColor = &HFFFF&
    End If
End Function

Function GDI_CreateSoildBrush(bColor As OLE_COLOR) As Long
    'Create a Brush form a picture handle
    GDI_CreateSoildBrush = CreateSolidBrush(GDI_TranslateColor(bColor))
End Function

Public Sub SetToolbarBG(hwnd As Long, hBmp As Long)
    'Set the toolbars background image
    DeleteObject SetClassLong(hwnd, GCL_HBRBACKGROUND, CreatePatternBrush(hBmp))
    InvalidateRect 0&, 0&, False
End Sub

Public Sub SetToolbarBK(hwnd As Long, hColor As OLE_COLOR)
    ' Set a toolbars Backcolor
    DeleteObject SetClassLong(hwnd, GCL_HBRBACKGROUND, GDI_CreateSoildBrush(hColor))
    InvalidateRect 0&, 0&, False
End Sub
