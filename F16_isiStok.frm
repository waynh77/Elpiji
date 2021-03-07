VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form F16_isiStok 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Persediaan Masuk/Isi Stok"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14820
   ClipControls    =   0   'False
   Icon            =   "F16_isiStok.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   14820
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   11640
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7680
      TabIndex        =   9
      Text            =   "Text9"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Text            =   "Text8"
      Top             =   2280
      Width           =   1050
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Text            =   "Text8"
      Top             =   1920
      Width           =   1050
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808000&
      Caption         =   "TIDAK"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   22
      Top             =   2280
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808000&
      Caption         =   "YA"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   21
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Text            =   "Text7"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Text            =   "Text6"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   7680
      TabIndex        =   10
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      Format          =   57212929
      CurrentDate     =   39688
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   7680
      TabIndex        =   7
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      Format          =   57212929
      CurrentDate     =   39688
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   7680
      TabIndex        =   23
      Text            =   "Text5"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "SUPPLIER BARU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4680
      MouseIcon       =   "F16_isiStok.frx":3482
      MousePointer    =   99  'Custom
      Picture         =   "F16_isiStok.frx":378C
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1560
      Width           =   1050
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   840
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "REFRESH"
      DownPicture     =   "F16_isiStok.frx":510E
      Height          =   855
      Index           =   5
      Left            =   12360
      MouseIcon       =   "F16_isiStok.frx":5878
      MousePointer    =   99  'Custom
      Picture         =   "F16_isiStok.frx":5B82
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      DownPicture     =   "F16_isiStok.frx":6C04
      Height          =   855
      Index           =   4
      Left            =   11160
      MouseIcon       =   "F16_isiStok.frx":736E
      MousePointer    =   99  'Custom
      Picture         =   "F16_isiStok.frx":7678
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CARI"
      DownPicture     =   "F16_isiStok.frx":86FA
      Height          =   855
      Index           =   3
      Left            =   9960
      MouseIcon       =   "F16_isiStok.frx":8E64
      MousePointer    =   99  'Custom
      Picture         =   "F16_isiStok.frx":916E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "HAPUS"
      DownPicture     =   "F16_isiStok.frx":A1F0
      Height          =   855
      Index           =   2
      Left            =   13080
      MouseIcon       =   "F16_isiStok.frx":A95A
      MousePointer    =   99  'Custom
      Picture         =   "F16_isiStok.frx":AC64
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "EDIT"
      DownPicture     =   "F16_isiStok.frx":BCE6
      Height          =   855
      Index           =   1
      Left            =   11520
      MouseIcon       =   "F16_isiStok.frx":C450
      MousePointer    =   99  'Custom
      Picture         =   "F16_isiStok.frx":C75A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "TAMBAH"
      DownPicture     =   "F16_isiStok.frx":D7DC
      Height          =   855
      Index           =   0
      Left            =   9960
      MouseIcon       =   "F16_isiStok.frx":DF46
      MousePointer    =   99  'Custom
      Picture         =   "F16_isiStok.frx":E250
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "KELUAR"
      DownPicture     =   "F16_isiStok.frx":F2D2
      Height          =   855
      Index           =   6
      Left            =   13560
      MouseIcon       =   "F16_isiStok.frx":FA3C
      MousePointer    =   99  'Custom
      Picture         =   "F16_isiStok.frx":FD46
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATA1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   3975
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   240
      TabIndex        =   25
      Top             =   2760
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   7858
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
      Left            =   2280
      Top             =   4560
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
            Picture         =   "F16_isiStok.frx":10DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F16_isiStok.frx":11E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F16_isiStok.frx":12EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F16_isiStok.frx":13F7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   7395
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20955
            Text            =   "PT Arjuna Wahana Putra - Sub Dealer Gas Elpiji Pertamina"
            TextSave        =   "PT Arjuna Wahana Putra - Sub Dealer Gas Elpiji Pertamina"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "3/1/2009"
            Key             =   "Tanggal Sekarang "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "10:18 PM"
            Key             =   "Jam Sekarang"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA ISI"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   15
      Left            =   6000
      TabIndex        =   42
      Top             =   840
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTY KOSONG"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   14
      Left            =   240
      TabIndex        =   41
      Top             =   2280
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTY ISI"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   40
      Top             =   1920
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS ISI"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   12
      Left            =   3120
      TabIndex        =   39
      Top             =   1560
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TABUNG+ISI"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   11
      Left            =   480
      TabIndex        =   38
      Top             =   3960
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA ISI"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   10
      Left            =   480
      TabIndex        =   37
      Top             =   3600
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DUE DATE"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   9
      Left            =   6000
      TabIndex        =   36
      Top             =   1920
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   8
      Left            =   6000
      TabIndex        =   35
      Top             =   120
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH BAYAR"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   7
      Left            =   6000
      TabIndex        =   34
      Top             =   1560
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA TABUNG"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH QTY"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   30
      Top             =   1560
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE SUPPLIER"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   29
      Top             =   1200
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA SUPPLIER"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   840
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA PRODUK"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   120
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE PRODUK"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HRG BELI TABUNG"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   3
      Left            =   6000
      TabIndex        =   24
      Top             =   480
      Width           =   1650
   End
End
Attribute VB_Name = "F16_isiStok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek As Boolean

Private Sub Combo1_Change()
isi_kode
isi_data1
isi_list
isi
End Sub

Sub isi_data1()
Data1.RecordSource = "select * from stok where kode_produk ='" & Text1 & "' order by kode_supplier asc , tanggal_masuk desc"
Data1.Refresh
End Sub

Sub isi_kode()
Data2.Refresh
If Not Data2.Recordset.BOF Then
    With Data2.Recordset
        .MoveFirst
        Do While Not .EOF
            If !nama_produk = Combo1 Then
                Text1 = !kode_produk
                .MoveLast
            End If
            .MoveNext
        Loop
    End With
    isi_harga
End If
End Sub

Sub isi_kode2()
Data3.Refresh
If Not Data3.Recordset.BOF Then
    With Data3.Recordset
        .MoveFirst
        Do While Not .EOF
            If !nama_supplier = Combo2 Then
                Text3 = !kode_supplier
                .MoveLast
            End If
            .MoveNext
        Loop
    End With
End If
End Sub

Sub isi_harga()
Dim cari As Boolean
'If Command1(0).Caption <> "SIMPAN" Then
Data4.Refresh
cari = False
With Data4.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until .EOF Or cari = True
        If Text1 = !kode_produk Then
            Text2(1) = Format(!harga_kosong, "###,###.00")
            Text6 = Format(!harga_isi, "###,###.00")
            Text7 = Format(!harga_kosong + !harga_isi, "###,###.00")
            cari = True
            .MovePrevious
        End If
        .MoveNext
    Loop
    If cari = False Then
        Text2(1) = 0
    End If
End If
End With
'End If
End Sub

Private Sub Combo1_Click()
isi_kode
If Combo2.Enabled = True Then
    Combo2.SetFocus
End If
isi_data1
isi_list
isi
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If KeyAscii = 13 And Combo2.Enabled = True Then
    Combo2.SetFocus
End If
End Sub

Private Sub Combo2_Change()
isi_kode2
End Sub

Private Sub Combo2_Click()
isi_kode2
If Text2(0).Enabled = True Then
    Text2(0).SetFocus
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If KeyAscii = 13 And DTPicker1.Enabled = True Then
    DTPicker1.SetFocus
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        tambah = True
        cmd_simpan
        buka
        kosong
        isi_cmb1
        isi_cmb2
        DTPicker1 = Date
        DTPicker2 = Date
        Combo1.SetFocus
    Else
        simpan
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
        If Not Data1.Recordset.BOF Then 'And Text2(1) <> 0 Then
            tambah = False
            cmd_simpan
            buka
            Text2(0) = Format(Text2(0), "###")
            Text9 = Format(Text9, "###")
            Text8(0).Enabled = True
            Text8(1).Enabled = True
            'Option1(0).Enabled = False
            'Option1(1).Enabled = False
            Combo1.Enabled = False
            Combo2.SetFocus
        Else
            MsgBox "Data Kosong/Belum Dipilih...", vbInformation, "Validasi Data"
        End If
    Else
        tutup
        Data1.Refresh
        Combo1.Enabled = True
        isi_list
        cmd_awal
        isi
    End If
Case 2
    With Data1.Recordset
        If Not .BOF And Text1 <> "" Then
            x = MsgBox("Apakah anda yakin ingin menghapus transaksi tanggal : " & Format(DTPicker1, "d mmmm yyyy") & " - Supplier(" & Combo2 & ")", vbYesNo, "Hapus Transaksi Produk " & Text1)
            If x = vbYes Then
                .Delete
                Data1.Refresh
                isi_list
                isi
            End If
        Else
            MsgBox "Data kosong/belum dipilih...", vbCritical, "Hapus Data"
        End If
    End With
Case 3
    cek = False
    cari = InputBox("Masukan Kode Produk", "Cari Data", "Kode Produk")
    If cari <> "Kode Produk" And cari <> "" Then
    With Data1.Recordset
        If Not .BOF Then
           Data1.Refresh
            .MoveFirst
            Do Until .EOF Or cek = True
                If !kode_produk Like ("*" & cari & "*") Then
                    cek = True
                End If
                .MoveNext
            Loop
            If cek = False Then
                MsgBox "Data tidak ditemukan", vbOKOnly, "Cari Data"
            Else
                MsgBox "Data ditemukan...", vbOKOnly, "Cari Data"
                .MovePrevious
                isi
                'Text1 = !kode_produk
                'Text2(0) = !harga_tabung
                'Text2(1) = !harga_kosong
                'cari_produk
            End If
        End If
    End With
    End If
Case 4
        CrystalReport1.ReportFileName = App.Path & "\Laporan Stok.rpt"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
Case 5
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
    Data4.Refresh
    isi
Case 6
    Unload Me
End Select
End Sub

Sub simpan()
If Text2(0) = "" Or Text4 = "" Then
    MsgBox "Data belum lengkap...", vbCritical, "Validasi Input"
    If Text2(0) = "" Then
        Text2(0).SetFocus
    Else
        Text4.SetFocus
    End If
Else
    If tambah = True Then
        cek = False
        'cek_data
        If cek = True Then
            MsgBox "Data Sudah Ada", vbInformation, "Validasi Data"
        Else
            With Data1.Recordset
                .AddNew
                transfer
                If Option1(0).Value = True Then
                    !qty_isi = Val(Text4)
                    !qty_kosong = 0
                Else
                    !qty_isi = 0
                    !qty_kosong = Val(Text4)
                End If
                !jam = Time
                tambah_beli
                .Update
                .MovePrevious
                tutup
                cmd_awal
                Data1.Refresh
                isi_list
                isi
                Combo1.Enabled = True
            End With
        End If
    Else
        If Val(Text4) = (Val(Text8(0)) + Val(Text8(1))) Then
            update_beli
            Data1.Recordset.Edit
            transfer
            Data1.Recordset!qty_isi = Val(Text8(0))
            Data1.Recordset!qty_kosong = Val(Text8(1))
            Data1.Recordset!jam = Time
            Data1.Recordset.Update
            tutup
            cmd_awal
            isi_list
            isi
            Combo1.Enabled = True
        Else
            MsgBox "Data quantity tidak balance...", vbCritical, "Perhitungan Salah"
            Text8(0).SetFocus
        End If
    End If
End If
End Sub

Sub tambah_beli()
With Data5.Recordset
    .AddNew
    !tgl_beli = DTPicker1
    !kode_supplier = Text3
    !nama_supplier = Combo2
    !qty_beli = Text4
    !harga_satuan = Val(Text2(0)) + Val(Text9)
    !kode_produk = Text1
    !nama_produk = Combo1
    If Option1(0).Value = True Then
        !status_produk = "TABUNG+ISI"
    Else
        !status_produk = "TABUNG KOSONG"
    End If
    !due_date = DTPicker2
    If DTPicker2 = Date Then
        !jumlah_bayar = !qty_beli * !harga_satuan
        !status_bayar = True
        !frek_bayar = 1
        'update bayar
        With Data6.Recordset
            .AddNew
            !tgl_byr = Date
            !jml_byr = Data5.Recordset!jumlah_bayar
            !keterangan = ""
            !tgl_beli = Date
            !kode_supplier = Text3
            !kode_produk = Text1
            !status_produk = Data5.Recordset!status_produk
            !jml_beli = Data5.Recordset!jumlah_bayar
            !qty = Text4
            !nama_supplier = Combo2
            !nama_produk = Combo1
            .Update
        End With
        'update kas
        With Data7.Recordset
            .AddNew
            !jenis_kas = "PENGELUARAN"
            !tgl_kas = Date
            !jam_kas = Time
            !nama_kas = "Beli " & Combo1
            !keterangan = Data6.Recordset!status_produk & " qty = " & Text4
            !nominal = Data6.Recordset!jml_byr
            .Update
        End With
    Else
        !jumlah_bayar = 0
        !status_bayar = False
        !frek_bayar = 0
        'isi remainder
        With Data8.Recordset
            .AddNew
            Call remain_auto
            !nomor = urut_remain
            !tgl = DTPicker2
            !waktu = "10:00"
            !Status = True
            !keterangan = "Hutang kepada " & Combo2 & " untuk Pembelian " & Combo1 & " " & Data5.Recordset!status_produk & " qty = " & Text4
            .Update
        End With
    End If
    .Update
End With
End Sub

Sub update_beli()
Data5.Refresh
With Data5.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
    If !tgl_beli = Data1.Recordset!tanggal_masuk And !kode_supplier = Data1.Recordset!kode_supplier And !kode_produk = Data1.Recordset!kode_produk And !qty_beli = Data1.Recordset!jumlah And !due_date = Data1.Recordset!due_date Then
            .Edit
            !tgl_beli = DTPicker1
            !kode_supplier = Text3
            !nama_supplier = Combo2
            !qty_beli = Text4
            !harga_satuan = Val(Text2(0)) + Val(Text9)
            !kode_produk = Text1
            !nama_produk = Combo1
            If Option1(0).Value = True Then
                !status_produk = "TABUNG+ISI"
            Else
                !status_produk = "TABUNG KOSONG"
            End If
            !due_date = DTPicker2
            !jumlah_bayar = 0
            !status_bayar = False
            !frek_bayar = 0
            .Update
        End If
    .MoveNext
    Loop
End If
End With
End Sub

Sub cek_data()
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until .EOF Or cek = True
        If Text1 = !kode_produk And Text3 = !kode_supplier Then
            cek = True
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub transfer()
With Data1.Recordset
    !kode_produk = Text1
    !kode_supplier = Text3
    !harga_tabung = Text2(0)
    !tanggal_masuk = DTPicker1
    !jumlah = Text4
    !due_date = DTPicker2
    If Option1(0).Value = True Then
        !Status = True
    Else
        !Status = False
    End If
    !harga_isi = Val(Text9)
End With
End Sub

Private Sub Command2_Click()
Me.Enabled = False
ctl2 = True
F04_Supplier.Show
End Sub

Private Sub DTPicker1_Click()
Text2(0).SetFocus
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2(0).Enabled = True Then
    Text2(0).SetFocus
End If
End Sub

Private Sub DTPicker2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1(0).SetFocus
End If
End Sub

Private Sub Form_Activate()
cek_stok
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh
DTPicker1 = Date
DTPicker2 = Date
isi_data1
isi_list
isi
isi_cmb1
isi_cmb2
End Sub

Private Sub Form_Load()
Call db_IsiStok
kosong
tutup
cmd_awal
End Sub

Private Sub Form_Unload(Cancel As Integer)
F08_TStok.Enabled = True
F08_TStok.Show
End Sub

Sub isi_list()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim gbr As ListImage
Dim jual As Double
Dim beli As Double
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
'Set head = ListView1.ColumnHeaders.Add(, , "Kode Produk", ListView1.Width / 10)
'Set head = ListView1.ColumnHeaders.Add(, , "Nama Produk", ListView1.Width / (10 / 2))
Set head = ListView1.ColumnHeaders.Add(, , "Tgl Masuk", ListView1.Width / 12)
Set head = ListView1.ColumnHeaders.Add(, , "Jam Masuk", ListView1.Width / 12, 2)
Set head = ListView1.ColumnHeaders.Add(, , "Jatuh Tempo", ListView1.Width / 12, 2)
Set head = ListView1.ColumnHeaders.Add(, , "Kode Supplier", ListView1.Width / 12, 2)
Set head = ListView1.ColumnHeaders.Add(, , "Nama Supplier", ListView1.Width / (12 / 2) - 100)
Set head = ListView1.ColumnHeaders.Add(, , "Qty Isi", ListView1.Width / 12, 2)
Set head = ListView1.ColumnHeaders.Add(, , "Qty Kosong", ListView1.Width / 12, 2)
Set head = ListView1.ColumnHeaders.Add(, , "Jumlah", ListView1.Width / 12, 2)
Set head = ListView1.ColumnHeaders.Add(, , "Harga Tabung", ListView1.Width / 12, 1)
Set head = ListView1.ColumnHeaders.Add(, , "Harga Isi", ListView1.Width / 12, 1)
Set head = ListView1.ColumnHeaders.Add(, , "Tabung+Isi", ListView1.Width / 12, 1)
'Set head = ListView1.ColumnHeaders.Add(, , "Status", ListView1.Width / 10, 2)
'Set head = ListView1.ColumnHeaders.Add(, , "Profit @", ListView1.Width / 11, 1)
ListView1.View = lvwReport
Data1.RecordSource = "select * from stok where kode_produk ='" & Text1 & "' order by kode_supplier asc , tanggal_masuk desc"
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
'        Set dtl = ListView1.ListItems.Add(, , !kode_produk)
'        Data2.Refresh
'        With Data2.Recordset
'            .MoveFirst
'            Do While Not .EOF
'                If !kode_produk = Data1.Recordset!kode_produk Then
'                    dtl.SubItems(1) = !nama_produk
'                    .MoveLast
'                End If
'                .MoveNext
'            Loop
'        End With
        Set dtl = ListView1.ListItems.Add(, , !tanggal_masuk)
'        dtl.SubItems(2) = !tanggal_masuk
        dtl.SubItems(1) = !jam
        dtl.SubItems(2) = !due_date
        dtl.SubItems(3) = !kode_supplier
        Data3.Refresh
        With Data3.Recordset
            .MoveFirst
            Do While Not .EOF
                If !kode_supplier = Data1.Recordset!kode_supplier Then
                    dtl.SubItems(4) = !nama_supplier
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        dtl.SubItems(5) = !qty_isi
        dtl.SubItems(6) = !qty_kosong
        dtl.SubItems(7) = !jumlah
        dtl.SubItems(8) = Format(!harga_tabung, "###,###,00")
        dtl.SubItems(9) = Format(!harga_isi, "###,###,00")
        dtl.SubItems(10) = Format(!harga_tabung + !harga_isi, "###,###,00")
'        If !Status = True Then
'            dtl.SubItems(8) = "Isi"
'        Else
'            dtl.SubItems(8) = "Kosong"
'        End If
        beli = !harga_tabung
        Data4.Refresh
        With Data4.Recordset
            If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                If !kode_produk = Data1.Recordset!kode_produk Then
                    'dtl.SubItems(8) = !harga_kosong
                    jual = !harga_kosong
                    .MoveLast
                End If
                .MoveNext
            Loop
            End If
        End With
        'dtl.SubItems(9) = Format(jual - beli, "###,###,00")
        .MoveNext
    Loop
End If
End With
Data1.Refresh
End Sub

Sub cari_produk()
Data2.Refresh
If Not Data2.Recordset.BOF Then
    With Data2.Recordset
        .MoveFirst
        Do While Not .EOF
            If !kode_produk = Text1 Then
                Combo1 = !nama_produk
                .MoveLast
            End If
            .MoveNext
        Loop
    End With
End If
End Sub

Sub cari_supplier()
Data3.Refresh
If Not Data3.Recordset.BOF Then
    With Data3.Recordset
        .MoveFirst
        Do While Not .EOF
            If !kode_supplier = Text3 Then
                Combo2 = !nama_supplier
                .MoveLast
            End If
            .MoveNext
        Loop
    End With
End If
End Sub

Sub kosong()
Text2(0) = ""
Text4 = ""
Combo1 = ""
Combo2 = ""
Text1 = ""
Text2(1) = ""
Text3 = ""
Text6 = ""
Text7 = ""
Text5 = ""
Text8(0) = ""
Text8(1) = ""
Text9 = ""
End Sub

Sub isi()
If Command1(0).Caption <> "SIMPAN" Then
With Data1.Recordset
    If Not .BOF Then
        Text3 = !kode_supplier
        Text1 = !kode_produk
        Text2(0) = Format(!harga_tabung, "###,###.00")
        Text9 = Format(!harga_isi, "###,###.00")
        Text4 = !jumlah
        Text5 = Format((!jumlah * !harga_tabung) + (!jumlah * !harga_isi), "###,###.00")
        DTPicker1 = !tanggal_masuk
        DTPicker2 = !due_date
        If !Status = True Then
            Option1(0).Value = True
        Else
            Option1(1).Value = True
        End If
        Text8(0) = !qty_isi
        Text8(1) = !qty_kosong
        cari_supplier
        cari_produk
    End If
End With
End If
End Sub

Sub tutup()
Text1.Enabled = False
Text2(0).Enabled = False
'Combo1.Enabled = False
Combo2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text2(1).Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
DTPicker1.Enabled = False
DTPicker2.Enabled = False
Text8(0).Enabled = False
Text8(1).Enabled = False
Text9.Enabled = False
End Sub

Sub buka()
Combo1.Enabled = True
Combo2.Enabled = True
Text2(0).Enabled = True
Text4.Enabled = True
DTPicker1.Enabled = True
DTPicker2.Enabled = True
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
Command2.Visible = True
Combo1.Enabled = True
Option1(0).Enabled = False
Option1(1).Enabled = False
ListView1.Enabled = True
Text9.Enabled = False
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
Command2.Visible = False
Combo1.Enabled = False
Option1(0).Enabled = True
Option1(1).Enabled = True
ListView1.Enabled = False
Text9.Enabled = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
cari_data
End Sub

Sub cari_data()
Dim beli As String
cek = False
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until cek = True Or .EOF
            beli = Format(!harga_tabung, "###,###")
            If !kode_produk = Text1 And !kode_supplier = ListView1.SelectedItem.ListSubItems(3).Text And !tanggal_masuk = ListView1.SelectedItem.Text And beli = ListView1.SelectedItem.ListSubItems(8).Text And !jumlah = ListView1.SelectedItem.ListSubItems(7).Text Then
                cek = True
                .MovePrevious
            End If
            .MoveNext
        Loop
        isi
    End If
End With
End Sub

Sub isi_cmb1()
Combo1.Clear
Data2.Refresh
With Data2.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo1.AddItem !nama_produk
            .MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
End With
End Sub

Sub isi_cmb2()
Combo2.Clear
Data3.Refresh
With Data3.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo2.AddItem !nama_supplier
            .MoveNext
        Loop
        Combo2.ListIndex = 0
    End If
End With
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Label1(15).Visible = True
    Text9.Visible = True
Case 1
    Label1(15).Visible = False
    Text9.Visible = False
    Text9 = 0
End Select
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Text9.Visible = True Then
            Text9.SetFocus
        Else
            Text4.SetFocus
        End If
    End If
Case 1
End Select
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Text5 = Format((Val(Text4) * Val(Text2(0))) + (Val(Text4) * Val(Text9)), "###,###.00")
        If Option1(0).Value = True Then
            Text8(0) = Val(Text4)
            Text8(1) = 0
        Else
            Text8(0) = 0
            Text8(1) = Val(Text4)
        End If
        Command1(0).SetFocus
    End If
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Sub cek_stok()
Data1.RecordSource = "stok"
Data1.Refresh
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !jumlah = 0 Then
                .Delete
            End If
            .MoveNext
        Loop
    End If
End With
End Sub
