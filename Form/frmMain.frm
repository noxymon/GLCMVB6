VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Haralick"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   719
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtW3 
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Text            =   "0"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtW2 
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Text            =   "0"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtW1 
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Text            =   "0"
      Top             =   2760
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog dialFile 
      Left            =   10800
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   6855
      Begin VB.PictureBox picUtama 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   0
         ScaleHeight     =   359
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   455
         TabIndex        =   7
         Top             =   0
         Width           =   6855
      End
   End
   Begin VB.CommandButton cmdGrayScale 
      Caption         =   "Gray Scale !"
      Height          =   735
      Left            =   7320
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   50
      Left            =   360
      SmallChange     =   10
      TabIndex        =   4
      Top             =   5520
      Width           =   6855
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5415
      LargeChange     =   50
      Left            =   120
      SmallChange     =   10
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   7320
      TabIndex        =   2
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract Haralick"
      Height          =   735
      Left            =   7320
      TabIndex        =   1
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   735
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "W3"
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "W2"
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "W1"
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FilePath As String

Private Sub LoadImage(Str As String)
    Dim pic As StdPicture
    Set pic = LoadPicture(Str)
    picUtama.AutoRedraw = True
    picUtama.Width = picUtama.ScaleX(pic.Width, vbHimetric, vbTwips)
    picUtama.Height = picUtama.ScaleY(pic.Height, vbHimetric, vbTwips)
    picUtama.PaintPicture pic, 0, 0, picUtama.ScaleX(pic.Width), picUtama.ScaleY(pic.Height)
    picUtama.Scale (0, 0)-(picUtama.ScaleX(pic.Width), picUtama.ScaleY(pic.Height))
    picUtama.Refresh
    Set pic = Nothing
End Sub

Private Sub cmdExtract_Click()
    Call GetTexture(picUtama, 0)
    Debug.Print ("ASM : " & GetASM)
    Debug.Print ("IDM : " & GetIDM)
    Debug.Print ("Contrast : " & GetContrast)
    Debug.Print ("Entrophy : " & GetEntropy)
    Debug.Print ("Correlation : " & GetCorrelation)
End Sub

Private Sub cmdGrayScale_Click()
    GrayScale picUtama, CDbl(txtW1.Text), CDbl(txtW2.Text), CDbl(txtW3.Text)
End Sub

Private Sub cmdOpen_Click()
    dialFile.FileName = vbNullString
    dialFile.Filter = "BMP | *.BMP| JPG | *.JPG"
    dialFile.ShowOpen
    If dialFile.FileName <> "" Then
        FilePath = dialFile.FileName
        picUtama.ScaleMode = 3
        picUtama.Picture = LoadPicture(FilePath)
        ReDim Values(0 To picUtama.ScaleWidth, 0 To picUtama.ScaleHeight) As Long
        'LoadImage dialFile.FileName
        VScroll1.Max = picUtama.ScaleHeight
        HScroll1.Max = picUtama.ScaleWidth
    Else
        FilePath = vbNullString
    End If
End Sub

Private Sub HScroll1_Change()
    picUtama.Left = HScroll1.Value * -1
End Sub

Private Sub VScroll1_Change()
    picUtama.Top = VScroll1.Value * -1
End Sub
