VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form F21_UlangNota 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Ulang Nota"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ClipControls    =   0   'False
   Icon            =   "F21_UlangNota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   840
      Width           =   2655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6800
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA PEMBELI"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   3480
      TabIndex        =   11
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TELP. PEMBELI"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMOR MEMBER"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMOR NOTA"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SILAHKAN MASUKAN KATA KUNCI UNTUK MENCARI NOTA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7725
   End
End
Attribute VB_Name = "F21_UlangNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Data1.RecordSource = "select * from nota where no_nota like '*" & Text1 & "*' and nO_member like '*" & Text2 & "*' and telp_pembeli like '*" & Text4 & "*' and nama_pembeli like '*" & Text3 & "*'"
    Data1.Refresh
    isi_list
    kosong
Case 1
    Dim ttl As Single
    Dim nama As String
    Dim jml As Double
    ttl = 0
    nama = ""
    If Not Data1.Recordset.BOF Then
        If Text1 <> "" And Text3 <> "" Then
            ctl5 = True
            F20_ctkNota.Data2.RecordSource = "select * from nota where no_nota ='" & ListView1.SelectedItem.Text & "'"
            F20_ctkNota.Data2.Refresh
            With F20_ctkNota
                .Data2.Recordset.MoveFirst
                .Label1(5).Caption = .Data2.Recordset!no_nota
                .Label1(6).Caption = "Banjarmasin, " & Format(.Data2.Recordset!tgl_nota, "d mmmm yyyy")
                .Label1(16).Caption = .Data2.Recordset!no_member
                .Label1(17).Caption = .Data2.Recordset!nama_pembeli
                .Label1(18).Caption = .Data2.Recordset!alamat_pembeli
                .Label1(19).Caption = .Data2.Recordset!telp_pembeli
                .Label1(21).Caption = "Petugas " & .Data2.Recordset!nama_petugas
                .Label1(22).Caption = "Lokasi " & .Data2.Recordset!kode_lokasi
                Do While Not .Data2.Recordset.EOF
                    .Data1.Refresh
                    .Data1.Recordset.MoveFirst
                    Do While Not .Data1.Recordset.EOF
                        If .Data2.Recordset!kode_produk = .Data1.Recordset!kode_produk Then
                            nama = .Data1.Recordset!nama_produk
                            .Data1.Recordset.MoveLast
                        End If
                        .Data1.Recordset.MoveNext
                    Loop
                    .ListBox1(0).AddItem .Data2.Recordset!status_produk & " - " & nama
                    .ListBox1(1).AddItem Format(.Data2.Recordset!nilai_transaksi, "###,###.00")
                    .ListBox1(2).AddItem "     " & .Data2.Recordset!qty
                    .ListBox1(3).AddItem Format(.Data2.Recordset!qty * .Data2.Recordset!nilai_transaksi, "###,###.00")
                    ttl = ttl + .Data2.Recordset!qty
                    jml = jml + .Data2.Recordset!nilai_transaksi * .Data2.Recordset!qty
                    .Data2.Recordset.MoveNext
                Loop
                .Text1(0) = ttl
                .Text1(1) = Format(jml, "###,###.00")
                .Show
            End With
            Unload Me
        Else
            MsgBox "Silahkan pilih data terlebih dahulu", vbInformation, "Pilih Data"
            ListView1.SetFocus
        End If
    Else
        MsgBox "Silahkan data masih kosong/tidak diketemukan", vbInformation, "Cari Data"
    End If
Case 2
    kosong
    Text1.SetFocus
End Select
End Sub

Private Sub Form_Activate()
'isi_list
End Sub

Private Sub Form_Load()
kosong
limiter
Call db_UlangNota
End Sub

Private Sub Form_Unload(Cancel As Integer)
F01_Main.Enabled = True
'F01_Main.Show
End Sub

Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Sub limiter()
Text1.MaxLength = 12
Text2.MaxLength = 10
Text3.MaxLength = 50
Text4.MaxLength = 20
End Sub

Sub isi_list()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim gbr As ListImage
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
Set head = ListView1.ColumnHeaders.Add(, , "No Nota", ListView1.Width / 6)
Set head = ListView1.ColumnHeaders.Add(, , "No Member", ListView1.Width / 6)
Set head = ListView1.ColumnHeaders.Add(, , "Nama Pembeli", ListView1.Width / (6 / 3))
Set head = ListView1.ColumnHeaders.Add(, , "Telp", ListView1.Width / 6)
ListView1.View = lvwReport
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1.ListItems.Add(, , !no_nota)
        dtl.SubItems(1) = !no_member
        dtl.SubItems(2) = !nama_pembeli
        dtl.SubItems(3) = !telp_pembeli
        .MoveNext
    Loop
    ListView1.SetFocus
'Else
'    MsgBox "Data tidak diketemukan...", vbInformation, "Tidak Ketemu"
End If
End With
Data1.Refresh
End Sub

Private Sub ListView1_Click()
If Not ListView1.ListItems.Count = 0 Then
isi
End If
End Sub

Sub isi()
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until !no_nota = ListView1.SelectedItem.Text
        .MoveNext
    Loop
    Text1 = !no_nota
    Text2 = !no_member
    Text3 = !nama_pembeli
    Text4 = !telp_pembeli
Else
    kosong
End If
End With
End Sub

Private Sub ListView1_GotFocus()
If Not ListView1.ListItems.Count = 0 Then
    isi
End If
End Sub

