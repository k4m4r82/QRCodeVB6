VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo QRCodeLib Menggunakan VB 6"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEncodeData 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "http://coding4ever.wordpress.com/"
      Top             =   3135
      Width           =   4935
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3540
      Width           =   1215
   End
   Begin VB.PictureBox picEncode 
      AutoRedraw      =   -1  'True
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEncode_Click()
    Dim qrEncoder As New QRCodeEncoder
    
    qrEncoder.QRCodeEncodeMode = ENCODE_MODE.ENCODE_MODE_BYTE
    qrEncoder.QRCodeScale = 4
    qrEncoder.QRCodeVersion = 7
    qrEncoder.QRCodeErrorCorrect = ERROR_CORRECTION.ERROR_CORRECTION_M
    
    picEncode.Picture = qrEncoder.EncodeVB6(txtEncodeData.Text)
    
End Sub

Private Sub cmdDecode_Click()
    Dim decoder As New QRCodeDecoder
    
    Dim decodedString As String
    
    Dim qrCodeImage As New QRCodeBitmapImage
    qrCodeImage.SetBitmap = picEncode.Picture
    
    decodedString = decoder.decodeVB6(qrCodeImage)
    
    MsgBox "Hasil decode : " & vbCrLf & decodedString, vbInformation, "Informasi"
End Sub
