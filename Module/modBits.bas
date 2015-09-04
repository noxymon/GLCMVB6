Attribute VB_Name = "Module2"
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

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Private bmpBits() As Byte
Private hBmp As Bitmap

Public ASM As Double
Public Contrast As Double
Public Correlation As Double
Public IDM As Double
Public Entropy As Double

Private Sub GetBits(pBox As PictureBox)
    Dim iRet As Long
    iRet = GetObject(pBox.Picture.Handle, Len(hBmp), hBmp)
    ReDim bmpBits(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    iRet = GetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, bmpBits(0, 0, 0))
End Sub

Private Sub SetBits(pBox As PictureBox)
    Dim iRet As Long
    iRet = SetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, bmpBits(0, 0, 0))
    Erase bmpBits
End Sub

Public Sub GrayScale(pic As PictureBox, ByVal BobotR As Double, ByVal BobotG As Double, ByVal BobotB As Double)
    Dim X As Long
    Dim Y As Long
    Dim R As Integer, G As Integer, b As Integer, Gray As Integer
    
    Call GetBits(pic)
    For X = 0 To hBmp.bmWidth - 1
        For Y = 0 To hBmp.bmHeight - 1
            R = bmpBits(0, X, Y)
            G = bmpBits(1, X, Y)
            b = bmpBits(2, X, Y)
            
            Gray = BobotR * R + BobotG * G + BobotB * b
            bmpBits(0, X, Y) = Gray
            bmpBits(1, X, Y) = Gray
            bmpBits(2, X, Y) = Gray
        Next Y
    Next X
    Call SetBits(pic)
    Call pic.Refresh
End Sub

Public Sub GetTexture(pic As PictureBox, degree As Integer)
    
    Dim glcm(257, 257) As Double
    Dim offset, i As Integer
    Dim a, b As Integer
    
    Dim counter As Long
    Dim R, R1 As Integer
    
    Dim X, Y As Integer
    Call GetBits(pic)
    If degree = 0 Then
        For Y = 0 To hBmp.bmHeight - 2
            offset = Y * hBmp.bmHeight
            For X = 0 To hBmp.bmWidth - 2
                R = bmpBits(0, X, Y)
                R1 = bmpBits(0, X + 1, Y)
                a = R And &HFF
                b = R1
                glcm(a, b) = glcm(a, b) + 1
                glcm(b, a) = glcm(b, a) + 1
                counter = counter + 2
            Next X
        Next Y
    ElseIf degree = 90 Then
        For Y = 0 To hBmp.bmHeight - 2
            offset = Y * hBmp.bmHeight
            For X = 0 To hBmp.bmWidth - 2
                R = bmpBits(0, X, Y)
                R1 = bmpBits(0, X, Y - 1)
                a = R
                b = R1
                glcm(a, b) = glcm(a, b) + 1
                glcm(b, a) = glcm(b, a) + 1
                counter = counter + 2
            Next X
        Next Y
    ElseIf degree = 180 Then
        For Y = 0 To hBmp.bmHeight - 2
            offset = Y * hBmp.bmHeight
            For X = 0 To hBmp.bmWidth - 2
                R = bmpBits(0, X, Y)
                R1 = bmpBits(0, X - 1, Y)
                a = R
                b = R1
                glcm(a, b) = glcm(a, b) + 1
                glcm(b, a) = glcm(b, a) + 1
                counter = counter + 2
            Next X
        Next Y
    ElseIf degree = 270 Then
        For Y = 0 To hBmp.bmHeight - 2
            offset = Y * hBmp.bmHeight
            For X = 0 To hBmp.bmWidth - 2
                R = bmpBits(0, X, Y)
                R1 = bmpBits(0, X, Y + 1)
                a = R
                b = R1
                glcm(a, b) = glcm(a, b) + 1
                glcm(b, a) = glcm(b, a) + 1
                counter = counter + 2
            Next X
        Next Y
    End If
    
    For a = 0 To 257
        For b = 0 To 257
            glcm(a, b) = glcm(a, b) / counter
        Next b
    Next a
    
    For a = 0 To 257
        For b = 0 To 257
            ASM = ASM + (glcm(a, b) * glcm(a, b))
        Next b
    Next a
    
    For a = 0 To 257
        For b = 0 To 257
            Contrast = Contrast + ((a - b) * (a - b) * glcm(a, b))
        Next b
    Next a
    
    Dim px, py As Double
    Dim stdevx, stdevy As Double
    
    For a = 0 To 257
        For b = 0 To 257
            px = px + (a * glcm(a, b))
            py = py + (b * glcm(a, b))
        Next b
    Next a
    
    For a = 0 To 257
        For b = 0 To 257
            stdevx = stdevx + (a - px) * (a - px) * glcm(a, b)
            stdevy = stdevy + (b - py) * (b - py) * glcm(a, b)
        Next b
    Next a
    
    For a = 0 To 257
        For b = 0 To 257
            Correlation = Correlation + ((a - px) * (b - py) * glcm(a, b) / (stdevx * stdevy))
        Next b
    Next a
    
    For a = 0 To 257
        For b = 0 To 257
            IDM = IDM + (glcm(a, b) / (1 + (a - b) * (a - b)))
        Next b
    Next a
    
    For a = 0 To 257
        For b = 0 To 257
            If glcm(a, b) <> 0 Then
                Entropy = Entropy - (glcm(a, b) * (Math.Log(glcm(a, b))))
            Else
                
            End If
        Next b
    Next a
End Sub

Public Function GetASM() As Double
    GetASM = ASM
End Function

Public Function GetContrast() As Double
    GetContrast = Contrast
End Function

Public Function GetCorrelation() As Double
    GetCorrelation = Correlation
End Function

Public Function GetIDM() As Double
    GetIDM = IDM
End Function

Public Function GetEntropy() As Double
    GetEntropy = Entropy
End Function
