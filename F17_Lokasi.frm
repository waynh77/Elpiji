VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form F17_Lokasi 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Lokasi"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   ClipControls    =   0   'False
   Icon            =   "F17_Lokasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2760
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   4320
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080FFFF&
      Height          =   885
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   480
      Width           =   4575
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATABASE PRODUK"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "KELUAR"
      DownPicture     =   "F17_Lokasi.frx":0CCA
      Height          =   735
      Index           =   6
      Left            =   5520
      MouseIcon       =   "F17_Lokasi.frx":1434
      MousePointer    =   99  'Custom
      Picture         =   "F17_Lokasi.frx":173E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "TAMBAH"
      DownPicture     =   "F17_Lokasi.frx":27C0
      Height          =   735
      Index           =   0
      Left            =   120
      MouseIcon       =   "F17_Lokasi.frx":2F2A
      MousePointer    =   99  'Custom
      Picture         =   "F17_Lokasi.frx":3234
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "EDIT"
      DownPicture     =   "F17_Lokasi.frx":42B6
      Height          =   735
      Index           =   1
      Left            =   1080
      MouseIcon       =   "F17_Lokasi.frx":4A20
      MousePointer    =   99  'Custom
      Picture         =   "F17_Lokasi.frx":4D2A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "HAPUS"
      DownPicture     =   "F17_Lokasi.frx":5DAC
      Height          =   735
      Index           =   2
      Left            =   2040
      MouseIcon       =   "F17_Lokasi.frx":6516
      MousePointer    =   99  'Custom
      Picture         =   "F17_Lokasi.frx":6820
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CARI"
      DownPicture     =   "F17_Lokasi.frx":78A2
      Height          =   735
      Index           =   3
      Left            =   3000
      MouseIcon       =   "F17_Lokasi.frx":800C
      MousePointer    =   99  'Custom
      Picture         =   "F17_Lokasi.frx":8316
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      DownPicture     =   "F17_Lokasi.frx":9398
      Height          =   735
      Index           =   4
      Left            =   3840
      MouseIcon       =   "F17_Lokasi.frx":9B02
      MousePointer    =   99  'Custom
      Picture         =   "F17_Lokasi.frx":9E0C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "REFRESH"
      DownPicture     =   "F17_Lokasi.frx":AE8E
      Height          =   735
      Index           =   5
      Left            =   4680
      MouseIcon       =   "F17_Lokasi.frx":B5F8
      MousePointer    =   99  'Custom
      Picture         =   "F17_Lokasi.frx":B902
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5775
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   10186
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648384
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F17_Lokasi.frx":C984
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F17_Lokasi.frx":DA16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F17_Lokasi.frx":EAA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F17_Lokasi.frx":FB3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BIAYA"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   13
      Top             =   120
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE LOKASI"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA LOKASI"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1650
   End
End
Attribute VB_Name = "F17_Lokasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek As Boolean

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        tambah = True
        cmd_simpan
        buka
        kosong
        Text1.SetFocus
    Else
        simpan
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
        If Not Data1.Recordset.BOF And Text1 <> "" And Text2 <> "" And Text3 <> "" Then
            tambah = False
            cmd_simpan
            buka
            Text1.SetFocus
            Text3 = Format(Text3, "###")
        Else
            MsgBox "Data Kosong/Belum Dipilih...", vbInformation, "Validasi Data"
        End If
    Else
        tutup
        Data1.Refresh
        isi
        isi_list
        cmd_awal
    End If
Case 2
        With Data1.Recordset
        If Not .BOF And Text2 <> "" Then
            x = MsgBox("Apakah anda yakin ingin menghapus Lokasi : " & Text1 & " - " & Text2, vbYesNo, "Hapus Data Lokasi " & Text2)
            If x = vbYes Then
                .Delete
                .MovePrevious
                If .BOF Then
                    Data1.Refresh
                    kosong
                End If
                isi_list
                isi
            End If
        Else
            MsgBox "Data kosong/belum dipilih...", vbCritical, "Hapus Data"
        End If
    End With
Case 3
    cek = False
    cari = InputBox("Masukan Nama Lokasi", "Cari Data", "Nama Lokasi")
    If cari <> "Nama Lokasi" And cari <> "" Then
    With Data1.Recordset
        If Not .BOF Then
           Data1.Refresh
            .MoveFirst
            Do Until .EOF Or cek = True
                If !nama_lokasi Like ("*" & cari & "*") Then
                    cek = True
                End If
                .MoveNext
            Loop
            If cek = False Then
                MsgBox "Data tidak ditemukan", vbOKOnly, "Cari Data"
            Else
                MsgBox "Data ditemukan...", vbOKOnly, "Cari Data"
                .MovePrevious
                Text1 = !kode_lokasi
                Text2 = !nama_lokasi
                Text3 = !biaya_kirim
            End If
        End If
    End With
    End If
Case 4
        CrystalReport1.ReportFileName = App.Path & "\Data Lokasi.rpt"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
Case 5
    Data1.Refresh
    isi
Case 6
    Unload Me
End Select
End Sub

Sub simpan()
If Text2 = "" Or Text1 = "" Or Text3 = "" Then
    MsgBox "Data belum lengkap...", vbCritical, "Validasi Input"
    If Text1 = "" Then
       Text1.SetFocus
    ElseIf Text2 = "" Then
        Text2.SetFocus
    ElseIf Text3 = "" Then
        Text3.SetFocus
    End If
Else
    If tambah = True Then
        cek = False
        cek_data
        If cek = True Then
            MsgBox "Data Sudah Ada", vbInformation, "Validasi Data"
        Else
            With Data1.Recordset
                .AddNew
                transfer
                .Update
                .MovePrevious
                tutup
                cmd_awal
                isi_list
            End With
        If ctl1 = True Then
            Unload Me
        End If
        End If
    Else
        Data1.Recordset.Edit
        transfer
        Data1.Recordset.Update
        tutup
        cmd_awal
        isi_list
    End If
End If
End Sub

Sub cek_data()
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Text1 = !kode_lokasi Then
            cek = True
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub transfer()
With Data1.Recordset
    !kode_lokasi = Text1
    !nama_lokasi = Text2
    !biaya_kirim = Text3
End With
End Sub

Private Sub Form_Activate()
Data1.Refresh
Text1.MaxLength = 3
Text2.MaxLength = 100
isi_list
isi
End Sub

Private Sub Form_Load()
Call db_lokasi
kosong
tutup
cmd_awal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    F06_Member.Enabled = True
    F06_Member.Show
End Sub

Sub isi_list()
Dim head As ColumnHeader
Dim dtl As ListItem
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
Set head = ListView1.ColumnHeaders.Add(, , "Kode Lokasi", ListView1.Width / 4)
Set head = ListView1.ColumnHeaders.Add(, , "Nama Lokasi", ListView1.Width / (4 / 2))
Set head = ListView1.ColumnHeaders.Add(, , "Biaya Kirim", ListView1.Width / (4), 1)
ListView1.View = lvwReport
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1.ListItems.Add(, , !kode_lokasi)
        dtl.SubItems(1) = !nama_lokasi
        dtl.SubItems(2) = Format(!biaya_kirim, "###,###.00")
        .MoveNext
    Loop
End If
End With
Data1.Refresh
End Sub

Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
End Sub

Sub isi()
If Command1(0).Caption <> "SIMPAN" Then
With Data1.Recordset
    If Not .BOF Then
        Text1 = !kode_lokasi
        Text2 = !nama_lokasi
        Text3 = Format(!biaya_kirim, "###,###.00")
    End If
End With
End If
End Sub

Sub tutup()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End Sub

Sub buka()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
End Sub

Sub cmd_awal()
Command1(0).Picture = ImageList1.ListImages(1).Picture
Command1(1).Picture = ImageList1.ListImages(3).Picture
Command1(0).Caption = "TAMBAH"
Command1(1).Caption = "EDIT"
Command1(2).Visible = True
Command1(3).Visible = True
Command1(4).Visible = True
Command1(5).Visible = True
Command1(6).Visible = True
ListView1.Enabled = True
End Sub

Sub cmd_simpan()
Command1(0).Picture = ImageList1.ListImages(2).Picture
Command1(1).Picture = ImageList1.ListImages(4).Picture
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Visible = False
Command1(3).Visible = False
Command1(4).Visible = False
Command1(5).Visible = False
Command1(6).Visible = False
ListView1.Enabled = False
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
cari_data
End Sub

Sub cari_data()
cek = False
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until !kode_lokasi = ListView1.SelectedItem.Text Or .EOF
            If !kode_lokasi = ListView1.SelectedItem.Text Then
                cek = True
            End If
            .MoveNext
        Loop
        isi
    End If
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
    Beep
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    Command1(0).SetFocus
End If
End Sub
