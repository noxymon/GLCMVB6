Attribute VB_Name = "Module1"
Option Explicit

Private Type Bitmap
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Sub GetPicture(pic As PictureBox, Path As String)
    Dim X As Long, Y As Long
    Dim holdBmp As Long, hMemDC As Long, PicInfo As Bitmap
    Dim sample As Long
    
    GetObject pic.Image, Len(PicInfo), PicInfo
    
    For X = 0 To PicInfo.bmWidth - 1
        For Y = 0 To PicInfo.bmHeight - 1
            sample = GetPixel(pic.hDC, X, Y)
            sample = (sample And &HFF) * &H10101
            SetPixel pic.hDC, X, Y, sample
        Next
    Next
    pic.Refresh
End Sub

Public Sub GetTexture(pic As PictureBox, degree As Integer)
    
    Dim glcm(257, 257) As Double
    Dim offset, i As Integer
    
    Dim X, Y As Integer
    If degree = 0 Then
        For Y = 0 To pic.Height
            offset = Y * pic.Width
            For X = 0 To pic.Width
                i = offset + X
            Next X
        Next Y
    End If
                
End Sub

