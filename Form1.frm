VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencari Nilai Maksimum"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim arrData(10) As Integer
Dim Max As Integer, i As Integer
  'Isi elemen array arrData
  arrData(0) = 12
  arrData(1) = 500
  arrData(2) = 92
  arrData(3) = 262
  arrData(4) = 112
  arrData(5) = 152
  arrData(6) = 887
  arrData(7) = 10
  arrData(8) = 120
  arrData(9) = 12
  'Inisialisasi variabel Max
  Max = 0
  'Bersihkan form
  Form1.Cls
  'Periksa semua isi array
  For i = 0 To 9
    'Cetak data-nya ke layar
    Print arrData(i)
    'Jika array indeks ke-i lebih besar dari Max
    If arrData(i) > Max Then
       'Tampung nilai Max
       Max = arrData(i)
    Else 'Jika tidak...
       'Nilai Max masih tetap yang sebelumnya
       Max = Max
    End If 'Akhir pemeriksaan isi array
  Next i
  'Tampilkan nilai maksimal setelah selesai iterasi
  MsgBox "Nilai maksimum = " & Max, _
         vbInformation, "Maksimum"
End Sub


