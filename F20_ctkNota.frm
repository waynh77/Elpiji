VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form F20_ctkNota 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Nota"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ClipControls    =   0   'False
   Icon            =   "F20_ctkNota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "F20_ctkNota.frx":0CCA
      Top             =   5760
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   6960
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      Height          =   855
      Left            =   7320
      Picture         =   "F20_ctkNota.frx":0D1E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   23
      Left            =   4680
      TabIndex        =   31
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Lokasi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   22
      Left            =   240
      TabIndex        =   29
      Top             =   5280
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Petugas Antar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   21
      Left            =   240
      TabIndex        =   28
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL  ===>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   20
      Left            =   4635
      TabIndex        =   27
      Top             =   5040
      Width           =   1245
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   2535
      Index           =   3
      Left            =   6960
      TabIndex        =   24
      Top             =   2400
      Width           =   1695
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;3858"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   2535
      Index           =   2
      Left            =   6000
      TabIndex        =   23
      Top             =   2400
      Width           =   2055
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3625;4471"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   2535
      Index           =   1
      Left            =   4320
      TabIndex        =   22
      Top             =   2400
      Width           =   1695
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;3858"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   19
      Left            =   6360
      TabIndex        =   21
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   18
      Left            =   6360
      TabIndex        =   20
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   6360
      TabIndex        =   19
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   6360
      TabIndex        =   18
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   15
      Left            =   6960
      TabIndex        =   17
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA @"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   14
      Left            =   4320
      TabIndex        =   16
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   13
      Left            =   5880
      TabIndex        =   15
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KETERANGAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   12
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   5745
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   2535
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   4215
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7435;4471"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   5040
      TabIndex        =   12
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   5040
      TabIndex        =   11
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   5040
      TabIndex        =   10
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Member"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   5040
      TabIndex        =   9
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kepada Yth,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   5040
      TabIndex        =   8
      Top             =   360
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banjarmasin,................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Width           =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No.Nota"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0511)335 1050 - 901 2828 - 901 12288"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jl. Lambung Mangkurat No.10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUB DEALER GAS ELPIJI PERTAMINA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PT. ARJUNA WAHANA PUTRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOKO ""EDISON"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "F20_ctkNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Visible = False
Me.PrintForm
If ctl5 = False Then
    simpan_trans
    hit_UlangStok
    update_penjualan
End If
Unload Me
Unload F09_Nota
End Sub

Private Sub Form_Activate()
If ctl5 = False Then
    isi_list
End If
End Sub

Sub simpan_trans()
With F09_Nota
    .dt_list.Refresh
    If Not .dt_list.Recordset.BOF Then
        .dt_list.Recordset.MoveFirst
        Do While Not .dt_list.Recordset.EOF And Not .dt_list.Recordset.BOF
            .dt_nota.Recordset.AddNew
            .dt_nota.Recordset!no_nota = .Caption
            .dt_nota.Recordset!tgl_nota = Date
            .dt_nota.Recordset!jam_nota = Time
            If .Combo1 <> "" Then
                .dt_nota.Recordset!status_member = True
            Else
                .dt_nota.Recordset!status_member = False
            End If
            .dt_nota.Recordset!no_member = .Combo1
            .dt_nota.Recordset!nama_pembeli = .Text1
            .dt_nota.Recordset!telp_pembeli = .Text3
            .dt_nota.Recordset!alamat_pembeli = .Text2
            .dt_nota.Recordset!kode_produk = .dt_list.Recordset!kode_produk
            .dt_produk.Refresh
            If Not .dt_produk.Recordset.BOF Then
                .dt_produk.Recordset.MoveFirst
                Do While Not .dt_produk.Recordset.EOF
                    If .dt_list.Recordset!kode_produk = .dt_produk.Recordset!kode_produk Then
                        .dt_nota.Recordset!nama_produk = .dt_produk.Recordset!nama_produk
                        .dt_produk.Recordset.MoveLast
                    End If
                    .dt_produk.Recordset.MoveNext
                Loop
            End If
            .dt_nota.Recordset!status_produk = .dt_list.Recordset!status_beli
            .dt_nota.Recordset!qty = .dt_list.Recordset!qty_jual
            .dt_nota.Recordset!nilai_transaksi = .dt_list.Recordset!harga_satuan
            .dt_nota.Recordset!harga_beli = .dt_list.Recordset!harga_beli
            .dt_nota.Recordset!kode_petugas = .Text16
            .dt_nota.Recordset!nama_petugas = .Combo2
            .dt_nota.Recordset!kode_lokasi = .Text4
            .dt_nota.Recordset!due_date = .DTPicker1
            .dt_nota.Recordset.Update
            update_penjualan
            .dt_list.Recordset.Delete
            .dt_list.Recordset.MoveNext
        Loop
    End If
End With
End Sub

Sub hit_UlangStok()
F09_Nota.dt_stok.RecordSource = "select * from stok"
F09_Nota.dt_stok.Refresh
With F09_Nota.dt_stok.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If (!qty_isi + !qty_kosong) <> !jumlah Then
                .Edit
                !jumlah = !qty_isi + !qty_kosong
                .Update
            End If
            .MoveNext
        Loop
    End If
End With
End Sub

Private Sub Form_Load()
Call db_ctknota
With F09_Nota
Label1(6).Caption = "Banjarmasin, " & Format(Date, "d mmmm yyyy")
Label1(16).Caption = ": " & .Combo1
Label1(17).Caption = ": " & .Text1
Label1(18).Caption = ": " & .Text2
Label1(19).Caption = ": " & .Text3
Label1(5).Caption = .Caption
Label1(21).Caption = "Petugas : " & .Combo2
Label1(22).Caption = "Lokasi " & .Text4
Text1(1) = Format(.Text8(2), "###,###.00")
If .DTPicker1 > Date Then
    Label1(23).Visible = True
    Label1(23).Caption = "Due Date : " & Format(.DTPicker1, "d mmmm yyyy")
Else
    Label1(23).Visible = False
End If
End With
End Sub

Sub isi_list()
Dim ttl As Single
Dim nama As String
ttl = 0
nama = ""
With F09_Nota.dt_list.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Data1.Refresh
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            If !kode_produk = Data1.Recordset!kode_produk Then
                nama = Data1.Recordset!nama_produk
                Data1.Recordset.MoveLast
            End If
            Data1.Recordset.MoveNext
        Loop
        ListBox1(0).AddItem !status_beli & " - " & nama
        ListBox1(1).AddItem Format(!harga_satuan, "###,###.00")
        ListBox1(2).AddItem "     " & !qty_jual
        ListBox1(3).AddItem Format(!jumlah, "###,###.00")
        ttl = ttl + !qty_jual
        .MoveNext
    Loop
End If
End With
Text1(0) = ttl
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ctl5 = False Then
    F09_Nota.Enabled = True
    F09_Nota.Show
Else
    F01_Main.Enabled = True
    F01_Main.Show
End If
ctl5 = False
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
keyacii = 0
End Sub

Sub update_penjualan()
If Not F09_Nota.dt_list.Recordset.BOF And Not F09_Nota.dt_list.Recordset.EOF Then
    With Data3.Recordset
        .AddNew
        !tgl_jual = Date
        !no_member = Mid(Label1(16).Caption, 3, 10)
        !nama_pembeli = Mid(Label1(17).Caption, 3)
        !qty_jual = F09_Nota.dt_list.Recordset!qty_jual
        !harga_satuan = F09_Nota.dt_list.Recordset!harga_satuan
        !kode_produk = F09_Nota.dt_list.Recordset!kode_produk
        Data1.Refresh
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            If F09_Nota.dt_list.Recordset!kode_produk = Data1.Recordset!kode_produk Then
                !nama_produk = Data1.Recordset!nama_produk
                Data1.Recordset.MoveLast
            End If
            Data1.Recordset.MoveNext
        Loop
        !status_produk = F09_Nota.dt_list.Recordset!status_beli
        !due_date = F09_Nota.DTPicker1
        If F09_Nota.DTPicker1 = Date Then
            !jumlah_bayar = !qty_jual * !harga_satuan
            !status_bayar = True
            !frek_bayar = 1
            With Data4.Recordset
                .AddNew
                !tgl_byr = Date
                !jml_byr = Data3.Recordset!jumlah_bayar
                !keterangan = ""
                !tgl_jual = Date
                !no_member = Data3.Recordset!no_member
                !kode_produk = Data3.Recordset!kode_produk
                !status_produk = Data3.Recordset!status_produk
                !jml_jual = !jml_byr
                !qty = Data3.Recordset!qty_jual
                !nama_member = Data3.Recordset!nama_pembeli
                !nama_produk = Data3.Recordset!nama_produk
                !no_nota = Label1(5).Caption
                .Update
            End With
            'update kas
            With Data5.Recordset
                .AddNew
                !jenis_kas = "PENERIMAAN"
                !tgl_kas = Date
                !jam_kas = Time
                !nama_kas = "Jual " & Data3.Recordset!nama_produk & " Nota : " & Label1(5).Caption
                !keterangan = Data3.Recordset!status_produk & " qty = " & Data3.Recordset!qty_jual
                !nominal = Data3.Recordset!jumlah_bayar
                .Update
            End With
        Else
            !jumlah_bayar = 0
            !status_bayar = False
            !frek_bayar = 0
        'isi remainder
        With Data6.Recordset
            .AddNew
            Call remain_auto
            !nomor = urut_remain
            !tgl = F09_Nota.DTPicker1
            !waktu = "10:00"
            !Status = True
            !keterangan = "Piutang " & Label1(5).Caption & " a/n " & Data3.Recordset!nama_pembeli & " Penjualan " & Data3.Recordset!nama_produk & " " & Data3.Recordset!status_produk & " qty = " & Data3.Recordset!qty_jual
            .Update
        End With
        End If
        !no_nota = Label1(5)
        .Update
    End With
End If
End Sub
