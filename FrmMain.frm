VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Mencari Luas & Keliling Bangun Datar"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox input4 
      Height          =   285
      Left            =   5640
      TabIndex        =   23
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox input3 
      Height          =   285
      Left            =   5640
      TabIndex        =   22
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox input2 
      Height          =   285
      Left            =   5640
      TabIndex        =   21
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton btnHasil 
      Caption         =   "Hasil"
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox input1 
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox operasi 
      Height          =   315
      ItemData        =   "FrmMain.frx":0000
      Left            =   1560
      List            =   "FrmMain.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton segilima 
      Caption         =   "Segilima Beraturan"
      Height          =   195
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton oval 
      Caption         =   "Elips"
      Height          =   195
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.OptionButton lingkaran 
      Caption         =   "Lingkaran"
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.OptionButton segitiga 
      Caption         =   "Segitiga"
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton layang 
      Caption         =   "Layang-Layang"
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.OptionButton trapesium 
      Caption         =   "Trapesium"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.OptionButton jajargenjang 
      Caption         =   "Jajargenjang"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.OptionButton belah 
      Caption         =   "Belah Ketupat"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.OptionButton persegiPjng 
      Caption         =   "Persegi Panjang"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame FrameImg 
      Caption         =   "Rumus dan Petunjuk"
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   7335
      Begin VB.Image img 
         Height          =   3135
         Left            =   960
         Picture         =   "FrmMain.frx":001E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.OptionButton persegi 
      Caption         =   "Persegi"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Line Line3 
      X1              =   3720
      X2              =   240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblHasil 
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblInput4 
      Caption         =   "lblInput4"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblInput3 
      Caption         =   "lblInput3"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblInput2 
      Caption         =   "lblInput2"
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblInput1 
      Caption         =   "lblInput1"
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   3840
      X2              =   3840
      Y1              =   720
      Y2              =   3120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Operasi :"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   1200
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Pilih Bangun Datar"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1920
      Y1              =   1320
      Y2              =   3120
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub belah_Click()
img.Picture = LoadPicture("img/belah.jpg")
input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""

If operasi.Text = "Luas" Then
    lblInput1.Caption = "Diagonal 1 :"
    lblInput2.Caption = "Diagonal 2 :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = False
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = False
    input4.Visible = False
Else
    lblInput1.Caption = "Panjang Sisi :"
    lblInput1.Visible = True
    lblInput2.Visible = False
    lblInput3.Visible = False
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = False
    input3.Visible = False
    input4.Visible = False
End If
End Sub

Private Sub btnHasil_Click()
If operasi.Text = "Luas" Then 'jika menghitung luas
    If persegi.Value = True Then
        hasil = Val(input1.Text) * Val(input1.Text)
    ElseIf persegiPjng.Value = True Then
        hasil = Val(input1.Text) * Val(input2.Text)
    ElseIf belah.Value = True Then
        hasil = Val(input1.Text) * Val(input2.Text) / 2
    ElseIf jajargenjang.Value = True Then
        hasil = Val(input1.Text) * Val(input2.Text) / 2
    ElseIf trapesium.Value = True Then
        hasil = (Val(input1.Text) + Val(input2.Text)) * Val(input3.Text) / 2
    ElseIf layang.Value = True Then
        hasil = Val(input1.Text) * Val(input2.Text) / 2
    ElseIf segitiga.Value = True Then
        hasil = Val(input1.Text) * Val(input2.Text) / 2
    ElseIf lingkaran.Value = True Then
        hasil = Val(input1.Text) * Val(input1.Text) * 22 / 7
    ElseIf oval.Value = True Then
        hasil = Val(input1.Text) / 2 * Val(input2.Text) / 2 * 22 / 7
    ElseIf segilima.Value = True Then
        hasil = Val(input1.Text) * Val(input1.Text) * 1.72
    Else
        lblHasil.Caption = "Bangun Datar Tidak Dikenali"
    End If
    hasil = Round(CDbl(hasil), 5) 'CDbl mengubah ke tipe data double, lalu menggunakan fungsi Round
    lblHasil.Caption = "Luas = " & hasil    'menampilkan hasil
Else    'jika menghitung keliling
    If persegi.Value = True Then
        hasil = Val(input1.Text) * 4
    ElseIf persegiPjng.Value = True Then
        hasil = (2 * Val(input1.Text)) + (2 * Val(input2.Text))
    ElseIf belah.Value = True Then
        hasil = Val(input1.Text) * 4
    ElseIf jajargenjang.Value = True Then
        hasil = (2 * Val(input1.Text)) + (2 * Val(input2.Text))
    ElseIf trapesium.Value = True Then
        hasil = Val(input1.Text) + Val(input2.Text) + Val(input3.Text) + Val(input4.Text)
    ElseIf layang.Value = True Then
        hasil = (2 * Val(input1.Text)) + (2 * Val(input2.Text))
    ElseIf segitiga.Value = True Then
        hasil = Val(input1.Text) + Val(input2.Text) + Val(input3.Text)
    ElseIf lingkaran.Value = True Then
        hasil = 2 * Val(input1.Text) * 22 / 7
    ElseIf oval.Value = True Then
        hasil = (Val(input1.Text) + Val(input2.Text)) / 2 * 22 / 7
    ElseIf segilima.Value = True Then
        hasil = Val(input1.Text) * 5
    Else
        lblHasil.Caption = "Bangun Datar Tidak Dikenali"
    End If
    hasil = Round(CDbl(hasil), 5) 'CDbl mengubah ke tipe data double, lalu menggunakan fungsi Round
    lblHasil.Caption = "Keliling = " & hasil    'menampilkan hasil
End If
End Sub

Function Round(nValue As Double, nDigits As Integer) As Double
'Fungsi untuk pembulatan dengan parameter nilai dan jumlah digit
Round = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
End Function

Private Sub Form_Load()
operasi.Text = "Luas"

img.Picture = LoadPicture("img/persegi.jpg")

lblInput1.Caption = "Panjang Sisi :"
lblInput1.Visible = True
lblInput2.Visible = False
lblInput3.Visible = False
lblInput4.Visible = False
input1.Visible = True
input2.Visible = False
input3.Visible = False
input4.Visible = False

End Sub

Private Sub jajargenjang_Click()
img.Picture = LoadPicture("img/jajargenjang.jpg")
input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""

If operasi.Text = "Luas" Then
    lblInput1.Caption = "Alas    :"
    lblInput2.Caption = "Tinggi  :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = False
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = False
    input4.Visible = False
Else
    lblInput1.Caption = "Sisi P :"
    lblInput2.Caption = "Sisi L :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = False
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = False
    input4.Visible = False
End If
End Sub

Private Sub layang_Click()
img.Picture = LoadPicture("img/layang.jpg")
input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""

If operasi.Text = "Luas" Then
    lblInput1.Caption = "Diagonal 1 :"
    lblInput2.Caption = "Diagonal 2 :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = False
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = False
    input4.Visible = False
Else
    lblInput1.Caption = "Sisi P :"
    lblInput2.Caption = "Sisi L :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = False
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = False
    input4.Visible = False
End If
End Sub

Private Sub lingkaran_Click()
img.Picture = LoadPicture("img/lingkaran.jpg")
input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""

lblInput1.Caption = "Jari-jari :"
lblInput1.Visible = True
lblInput2.Visible = False
lblInput3.Visible = False
lblInput4.Visible = False

input1.Visible = True
input2.Visible = False
input3.Visible = False
input4.Visible = False
End Sub

Private Sub oval_Click()
img.Picture = LoadPicture("img/elip.jpg")
lblInput1.Caption = "Diameter A :"
lblInput2.Caption = "Diameter B :"
lblInput1.Visible = True
lblInput2.Visible = True
lblInput3.Visible = False
lblInput4.Visible = False

input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""
input1.Visible = True
input2.Visible = True
input3.Visible = False
input4.Visible = False
End Sub

Private Sub persegi_Click()
lblInput1.Caption = "Panjang Sisi :"
lblInput1.Visible = True
lblInput2.Visible = False
lblInput3.Visible = False
lblInput4.Visible = False

input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""
input1.Visible = True
input2.Visible = False
input3.Visible = False
input4.Visible = False
End Sub

Private Sub persegiPjng_Click()
img.Picture = LoadPicture("img/panjang.jpg")
input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""

If operasi.Text = "Luas" Then
    lblInput1.Caption = "Panjang :"
    lblInput2.Caption = "Lebar   :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = False
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = False
    input4.Visible = False
Else
    lblInput1.Caption = "Panjang :"
    lblInput2.Caption = "Lebar   :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = False
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = False
    input4.Visible = False
End If

End Sub

Private Sub segilima_Click()
img.Picture = LoadPicture("img/segilima.jpg")
lblInput1.Caption = "Panjang Sisi :"
lblInput1.Visible = True
lblInput2.Visible = False
lblInput3.Visible = False
lblInput4.Visible = False

input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""
input1.Visible = True
input2.Visible = False
input3.Visible = False
input4.Visible = False
End Sub

Private Sub segitiga_Click()
img.Picture = LoadPicture("img/segitiga.jpg")
input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""

If operasi.Text = "Luas" Then
    lblInput1.Caption = "Alas   :"
    lblInput2.Caption = "Tinggi :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = False
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = False
    input4.Visible = False
Else
    lblInput1.Caption = "Sisi 1 :"
    lblInput2.Caption = "Sisi 2 :"
    lblInput3.Caption = "Sisi 3 :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = True
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = True
    input4.Visible = False
End If
End Sub

Private Sub trapesium_Click()
img.Picture = LoadPicture("img/trapesium.jpg")
input1.SetFocus
input1.Text = ""
input2.Text = ""
input3.Text = ""
input4.Text = ""

If operasi.Text = "Luas" Then
    lblInput1.Caption = "Sisi A :"
    lblInput2.Caption = "Sisi B :"
    lblInput3.Caption = "Tinggi :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = True
    lblInput4.Visible = False

    input1.Visible = True
    input2.Visible = True
    input3.Visible = True
    input4.Visible = False
Else
    lblInput1.Caption = "Sisi 1 :"
    lblInput2.Caption = "Sisi 2 :"
    lblInput3.Caption = "Sisi 3 :"
    lblInput4.Caption = "Sisi 4 :"
    lblInput1.Visible = True
    lblInput2.Visible = True
    lblInput3.Visible = True
    lblInput4.Visible = True

    input1.Visible = True
    input2.Visible = True
    input3.Visible = True
    input4.Visible = True
End If
End Sub
