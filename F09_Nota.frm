VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form F09_Nota 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Penjualan"
   ClientHeight    =   9900
   ClientLeft      =   1725
   ClientTop       =   1455
   ClientWidth     =   12180
   ClipControls    =   0   'False
   Icon            =   "F09_Nota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   12180
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Text            =   "Combo4"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ARJUNA WAHANA PUTRA"
      Height          =   315
      Left            =   240
      TabIndex        =   64
      Top             =   240
      Width           =   5655
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6360
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "NEW"
      Height          =   315
      Index           =   1
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "SEARCH"
      Height          =   315
      Index           =   0
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   240
      Width           =   1095
   End
   Begin VB.Data dt_nota 
      Caption         =   "dt_nota"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dt_list 
      Caption         =   "dt_list"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dt_lokasi 
      Caption         =   "dt_lokasi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   9720
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   19595265
      CurrentDate     =   39697
   End
   Begin VB.Data dt_stok 
      Caption         =   "dt_stok"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      ForeColor       =   &H008080FF&
      Height          =   735
      Index           =   1
      Left            =   6120
      TabIndex        =   56
      Top             =   7200
      Width           =   5895
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "YA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "TIDAK"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BELI TABUNG+ISI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   285
         Index           =   21
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   2385
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      ForeColor       =   &H008080FF&
      Height          =   735
      Index           =   0
      Left            =   6120
      TabIndex        =   54
      Top             =   6360
      Width           =   5895
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "TIDAK"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "YA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ISI ULANG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   285
         Index           =   20
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   2385
      End
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Height          =   285
      Left            =   8880
      TabIndex        =   16
      Text            =   "Text15"
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   52
      Text            =   "Text14"
      Top             =   9960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   50
      Text            =   "Text13"
      Top             =   9480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   49
      Text            =   "Text12"
      Top             =   9120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   47
      Text            =   "Text11"
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   46
      Text            =   "Text10"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   45
      Text            =   "Text9"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   6360
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1920
      TabIndex        =   14
      Text            =   "Combo3"
      Top             =   6000
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6360
      Top             =   2280
   End
   Begin VB.Data dt_harga 
      Caption         =   "dt_harga"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dt_produk 
      Caption         =   "dt_produk"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dt_petugas 
      Caption         =   "dt_petugas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dt_member 
      Caption         =   "dt_member"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Index           =   2
      Left            =   9720
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   3720
      Width           =   3855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808000&
      Caption         =   "TIDAK"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   33
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808000&
      Caption         =   "YA"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   32
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "PREVIEW"
      DownPicture     =   "F09_Nota.frx":3482
      Height          =   735
      Index           =   3
      Left            =   3720
      MouseIcon       =   "F09_Nota.frx":3BEC
      MousePointer    =   99  'Custom
      Picture         =   "F09_Nota.frx":3EF6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "HAPUS"
      DownPicture     =   "F09_Nota.frx":4F78
      Height          =   735
      Index           =   2
      Left            =   2520
      MouseIcon       =   "F09_Nota.frx":56E2
      MousePointer    =   99  'Custom
      Picture         =   "F09_Nota.frx":59EC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "EDIT"
      DownPicture     =   "F09_Nota.frx":6A6E
      Height          =   735
      Index           =   1
      Left            =   1320
      MouseIcon       =   "F09_Nota.frx":71D8
      MousePointer    =   99  'Custom
      Picture         =   "F09_Nota.frx":74E2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "TAMBAH"
      DownPicture     =   "F09_Nota.frx":8564
      Height          =   735
      Index           =   0
      Left            =   120
      MouseIcon       =   "F09_Nota.frx":8CCE
      MousePointer    =   99  'Custom
      Picture         =   "F09_Nota.frx":8FD8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "KELUAR"
      DownPicture     =   "F09_Nota.frx":A05A
      Height          =   735
      Index           =   4
      Left            =   4920
      MouseIcon       =   "F09_Nota.frx":A7C4
      MousePointer    =   99  'Custom
      Picture         =   "F09_Nota.frx":AACE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "F09_Nota.frx":BB50
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   2040
      TabIndex        =   28
      Text            =   "Text4"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "F09_Nota.frx":BB56
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   2040
      TabIndex        =   27
      Text            =   "Combo1"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   21
      Text            =   "Text16"
      Top             =   8040
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   6120
      TabIndex        =   59
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7011
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   3480
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
            Picture         =   "F09_Nota.frx":BB5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F09_Nota.frx":CBEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F09_Nota.frx":DC80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F09_Nota.frx":ED12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   9720
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal,jan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   10710
      TabIndex        =   63
      Top             =   120
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FF80&
      Height          =   1095
      Left            =   120
      Top             =   6840
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL JATUH TEMPO ===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   23
      Left            =   6915
      TabIndex        =   60
      Top             =   5280
      Width           =   2790
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   120
      X2              =   12000
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA SATUAN @ Rp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Index           =   22
      Left            =   6120
      TabIndex        =   58
      Top             =   8040
      Width           =   4185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH QTY BELI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Index           =   19
      Left            =   6240
      TabIndex        =   53
      Top             =   6000
      Width           =   2505
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA ISI ULANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   18
      Left            =   120
      TabIndex        =   51
      Top             =   8880
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA JUAL SATUAN @"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   17
      Left            =   3240
      TabIndex        =   48
      Top             =   8760
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL STOK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   44
      Top             =   8040
      Width           =   4425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STOK KOSONG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   15
      Left            =   240
      TabIndex        =   43
      Top             =   7440
      Width           =   4305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STOK TABUNG ISI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   14
      Left            =   240
      TabIndex        =   42
      Top             =   7080
      Width           =   4305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE PRODUK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   13
      Left            =   120
      TabIndex        =   41
      Top             =   6360
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA PRODUK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   40
      Top             =   6000
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL TRANSAKSI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   6120
      TabIndex        =   36
      Top             =   4680
      Width           =   3570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BIAYA KIRIM/ANTAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   6120
      TabIndex        =   35
      Top             =   4200
      Width           =   3570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUB TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   6120
      TabIndex        =   34
      Top             =   3840
      Width           =   3570
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   1215
      Index           =   1
      Left            =   120
      Top             =   3480
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KIRIM/ANTAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   8
      Left            =   3480
      TabIndex        =   31
      Top             =   4440
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE PETUGAS"
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
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   30
      Top             =   4200
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA PETUGAS"
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
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   29
      Top             =   3720
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   3135
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AREA LOKASI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   26
      Top             =   2400
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE LOKASI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   25
      Top             =   2040
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TELP."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   24
      Top             =   1680
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALAMAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA MEMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   600
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NO MEMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1665
   End
End
Attribute VB_Name = "F09_Nota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim kd_lok As String
Dim cek As Boolean
Dim cek_dat As Boolean
Dim beli As Double

Private Sub Combo1_Change()
isi_member
hit_ttl
End Sub

Private Sub Combo1_Click()
isi_member
hit_ttl
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
isi_member
End Sub

Private Sub Combo2_Change()
isi_petugas
End Sub

Private Sub Combo2_Click()
isi_petugas
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_Change()
isi_produk
'hit_trans
End Sub

Private Sub Combo3_Click()
isi_produk
'hit_trans
Text15.SetFocus
End Sub

Private Sub Combo4_Change()
Text4 = Combo4
isi_lokasi
End Sub

Private Sub Combo4_Click()
Text4 = Combo4
isi_lokasi
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Dim x As String
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        cmd_simpan
        tambah = True
        kosong2
        isi_cmb3
        Text16 = ""
        Text15.SetFocus
    Else
        simpan
    End If
    Timer1.Enabled = True
Case 1
    If Command1(1).Caption = "EDIT" Then
        If Not dt_list.Recordset.BOF And Not dt_list.Recordset.EOF And Text15 <> "" Then
            cmd_simpan
            tambah = False
            Frame1(0).Enabled = False
            Frame1(1).Enabled = False
        Else
            MsgBox "Data masih kosong/belum dipilih...", vbInformation, "Validasi Data"
        End If
    Else
        Frame1(0).Enabled = True
        Frame1(1).Enabled = True
        cmd_awal
    End If
    Timer1.Enabled = True
Case 2
    If Not dt_list.Recordset.BOF And Not dt_list.Recordset.EOF Then
        x = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
        If x = vbYes Then
            edit_stok3
            dt_list.Recordset.Delete
            dt_list.Refresh
            isi_list
        End If
    Else
        MsgBox "Data masih kosong/belum dipilih...", vbInformation, "Validasi Data"
    End If
Case 3
    If dt_list.Recordset.BOF Then
        MsgBox "Belum ada transaksi,silahkan melakukan transaksi terlebih dahulu", vbInformation, "Validasi Data"
    Else
        Me.Enabled = False
        F20_ctkNota.Show
    End If
Case 4
    Unload Me
End Select
End Sub

Sub simpan()
If Text15 = "" Or Text16 = "" Then
    If Text15 = "" Then
        MsgBox "Jumlah Quantity pembelian belum di isi...", vbCritical, "Validasi Input"
        Text15.SetFocus
    ElseIf Text16 = "" Then
        MsgBox "Nilai transaksi pembelian belum di isi...", vbCritical, "Validasi Input"
        Text16.SetFocus
    End If
Else
    cek_stok
    If cek = False Then
        MsgBox "Jumlah Quantity tidak valid...", vbCritical, "Validasi Input"
        Text15.SetFocus
    Else
        With dt_list.Recordset
        If tambah = True Then
            cek_data
        Else
            cek_dat = False
        End If
        If cek_dat = True Then
            MsgBox "Data sudah ada, silahkan melakukan transaksi edit/isi dengan produk/status yang lain...", vbInformation, "Validasi Data"
        Else
            If tambah = True Then
                edit_stok
                .AddNew
            Else
                edit_stok2
                .Edit
            End If
            !kode_produk = Text7
            !qty_jual = Text15
            If Option2(0).Value = True Then
                !status_beli = "ISI ULANG"
            Else
                If Option2(3).Value = True Then
                    !status_beli = "TABUNG+ISI"
                Else
                    !status_beli = "TABUNG KOSONG"
                End If
            End If
            !harga_satuan = Format(Text16, "###")
            !jumlah = Val(Text15) * Format(Text16, "###")
            !harga_beli = beli
            .Update
            Frame1(0).Enabled = True
            Frame1(1).Enabled = True
            isi_list
            cmd_awal
        End If
        End With
    End If
End If
End Sub

Sub edit_stok()
Dim quanti As Single
Dim tot As Double
quanti = Val(Text15)
tot = Val(Text15)
beli = 0
With dt_stok.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until quanti = 0
        .Edit
        If Option2(0).Value = True Then
            If !qty_isi > 0 Then
                If quanti > !qty_isi Then
                    !qty_kosong = !qty_kosong + !qty_isi
                    beli = beli + !harga_isi * !qty_isi
                    quanti = quanti - !qty_isi
                    !qty_isi = 0
                Else
                    !qty_kosong = !qty_kosong + quanti
                    !qty_isi = !qty_isi - quanti
                    beli = !harga_isi * quanti
                    quanti = 0
                End If
            End If
        Else
            If Option2(3).Value = True Then
                If !qty_isi > 0 Then
                    If quanti > !qty_isi Then
                        '!jumlah = !jumlah - !qty_isi
                        quanti = quanti - !qty_isi
                        beli = beli + (!harga_tabung + !harga_isi) * !qty_isi
                        !qty_isi = 0
                    Else
                        !qty_isi = !qty_isi - quanti
                        '!jumlah = !jumlah - quanti
                         beli = (!harga_tabung + !harga_isi) * quanti
                       quanti = 0
                    End If
                End If
            Else
                If !qty_kosong > 0 Then
                    If quanti > !qty_kosong Then
                        '!jumlah = !jumlah - !Qty_kosong
                        quanti = quanti - !qty_kosong
                        beli = beli + !harga_tabung * !qty_kosong
                        !qty_kosong = 0
                    Else
                        !qty_kosong = !qty_kosong - quanti
                        '!jumlah = !jumlah - quanti
                        beli = !harga_tabung * quanti
                        quanti = 0
                    End If
                End If
            End If
        End If
        .Update
        .MoveNext
    Loop
    beli = beli / tot
End If
End With
dt_stok.Refresh
End Sub

Sub edit_stok2()
Dim quanti As Single
Dim tot As Double
beli = 0
quanti = Val(Text15) - dt_list.Recordset!qty_jual
tot = Val(Text15) - dt_list.Recordset!qty_jual
If quanti <> 0 Then
    With dt_stok.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until quanti = 0
            .Edit
            If Option2(0).Value = True Then
                If !qty_isi > 0 Then
                    If quanti > !qty_isi Then
                        !qty_kosong = !qty_kosong + !qty_isi
                        quanti = quanti - !qty_isi
                        beli = beli + !harga_isi * !qty_isi
                        !qty_isi = 0
                    Else
                        !qty_kosong = !qty_kosong + quanti
                        !qty_isi = !qty_isi - quanti
                        beli = beli + !harga_isi * quanti
                        quanti = 0
                    End If
                End If
            Else
                If Option2(3).Value = True Then
                    If !qty_isi > 0 Then
                        If quanti > !qty_isi Then
                            '!jumlah = !jumlah - !qty_isi
                            quanti = quanti - !qty_isi
                            beli = beli + (!harga_tabung + !harga_isi) * !qty_isi
                            !qty_isi = 0
                        Else
                            !qty_isi = !qty_isi - quanti
                            '!jumlah = !jumlah - quanti
                             beli = (!harga_tabung + !harga_isi) * quanti
                            quanti = 0
                        End If
                    End If
                Else
                    If !qty_kosong > 0 Then
                        If quanti > !qty_kosong Then
                            '!jumlah = !jumlah - !Qty_kosong
                            quanti = quanti - !qty_kosong
                            beli = beli + !harga_tabung * !qty_kosong
                            !qty_kosong = 0
                        Else
                            !qty_kosong = !qty_kosong - quanti
                            '!jumlah = !jumlah - quanti
                            beli = !harga_tabung * quanti
                            quanti = 0
                        End If
                    End If
                End If
            End If
            .Update
            .MoveNext
        Loop
        beli = ((dt_list.Recordset!harga_beli + (beli / tot))) / 2
    End If
    End With
End If
dt_stok.Refresh
End Sub

Sub edit_stok3()
Dim quanti As Single
quanti = dt_list.Recordset!qty_jual
dt_stok.Refresh
With dt_stok.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until quanti = 0 Or .EOF
        .Edit
        If Option2(0).Value = True Then
            If quanti + !qty_kosong > !jumlah Then
                !qty_isi = !qty_isi + !qty_kosong
                quanti = quanti + !qty_kosong - !jumlah
                !qty_kosong = 0
            Else
                !qty_kosong = !qty_kosong - quanti
                !qty_isi = !qty_isi + quanti
                quanti = 0
            End If
        Else
            If Option2(3).Value = True Then
                If quanti + !qty_kosong + !qty_isi > !jumlah Then
                    quanti = quanti - (jumlah - !qty_isi - !qty_kosong)
                    !qty_isi = !qty_isi + !jumlah - !qty_kosong
                Else
                    !qty_isi = !qty_isi + quanti
                    quanti = 0
                End If
            Else
                If quanti + !qty_kosong + !qty_isi > !jumlah Then
                    quanti = quanti - (jumlah - !qty_isi - !qty_kosong)
                    !qty_kosong = !qty_kosong + !jumlah - !qty_isi
                Else
                    !qty_kosong = !qty_kosong + quanti
                    quanti = 0
                End If
            End If
        End If
        .Update
        .MoveNext
    Loop
End If
End With
dt_stok.Refresh
End Sub

Sub edit_stok4()
Dim quanti As Single
dt_list.Refresh
With dt_list.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        quanti = dt_list.Recordset!qty_jual
        dt_stok.Refresh
        With dt_stok.Recordset
        If Not .BOF Then
            .MoveFirst
            Do Until quanti = 0 Or .EOF
                .Edit
                If dt_list.Recordset!status_beli = "ISI ULANG" Then
                    If quanti + !qty_kosong > !jumlah Then
                        !qty_isi = !qty_isi + !qty_kosong
                        quanti = quanti + !qty_kosong - !jumlah
                        !qty_kosong = 0
                    Else
                        !qty_kosong = !qty_kosong - quanti
                        !qty_isi = !qty_isi + quanti
                        quanti = 0
                    End If
                Else
                    If dt_list.Recordset!status_beli = "TABUNG+ISI" Then
                        If quanti + !qty_kosong + !qty_isi > !jumlah Then
                            quanti = quanti - (jumlah - !qty_isi - !qty_kosong)
                            !qty_isi = !qty_isi + !jumlah - !qty_kosong
                        Else
                            !qty_isi = !qty_isi + quanti
                            quanti = 0
                        End If
                    Else
                        If quanti + !qty_kosong + !qty_isi > !jumlah Then
                            quanti = quanti - (!jumlah - !qty_isi - !qty_kosong)
                            !qty_kosong = !qty_kosong + !jumlah - !qty_isi
                        Else
                            !qty_kosong = !qty_kosong + quanti
                            quanti = 0
                        End If
                    End If
                End If
                .Update
                .MoveNext
            Loop
        End If
        End With
        .Delete
        .MoveNext
    Loop
End If
End With
dt_stok.Refresh
End Sub

Sub cek_data()
Dim stat As String
cek_dat = False
If Option2(0).Value = True Then
    stat = "ISI ULANG"
Else
    If Option2(3).Value = True Then
        stat = "TABUNG+ISI"
    Else
        stat = "TABUNG KOSONG"
    End If
End If
With dt_list.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If !kode_produk = Text7 And !status_beli = stat Then
            cek_dat = True
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub isi_list()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim sub_ttl As Double
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
Set head = ListView1.ColumnHeaders.Add(, , "Nama Produk", ListView1.Width / (5 / 2))
Set head = ListView1.ColumnHeaders.Add(, , "Keterangan", ListView1.Width / (5 / 2) - 100)
Set head = ListView1.ColumnHeaders.Add(, , "QTY", ListView1.Width / 5, 2)
Set head = ListView1.ColumnHeaders.Add(, , "Harga @", ListView1.Width / 5, 1)
Set head = ListView1.ColumnHeaders.Add(, , "Jumlah", ListView1.Width / 5, 1)
ListView1.View = lvwReport
sub_ttl = 0
With dt_list.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        dt_produk.Refresh
        With dt_produk.Recordset
            .MoveFirst
            Do While Not .EOF
                If !kode_produk = dt_list.Recordset!kode_produk Then
                    Set dtl = ListView1.ListItems.Add(, , !nama_produk)
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        dtl.SubItems(1) = !status_beli
        dtl.SubItems(2) = !qty_jual
        dtl.SubItems(3) = Format(!harga_satuan, "###,###,00")
        dtl.SubItems(4) = Format(!jumlah, "###,###,00")
        sub_ttl = sub_ttl + !jumlah
        .MoveNext
    Loop
End If
End With
Text8(0) = Format(sub_ttl, "###,###.00")
hit_ttl
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0
    Me.Enabled = False
    F19_cariMember.Show
Case 1
    Me.Enabled = False
    F06_Member.Show
    ctl4 = True
End Select
End Sub

Private Sub Form_Activate()
Call db_Nota
isi_cmb3
isi_list
Text8(0) = Format(Text8(0), "###,###.00")
nota_auto
Label2(3).Caption = Format(Date, "d mmmm yyyy") & ", " & Format(Time, "hh:mm:ss")
End Sub

Private Sub Form_Load()
Me.Height = 6240
DTPicker1 = Date
Call db_Nota
Option1(0).Value = True
Option2(0).Value = True
Option2(3).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim a As String
dt_list.Refresh
If Not dt_list.Recordset.BOF Then
    a = MsgBox("Transaksi akan dihapus jika belum dicetak, apakah anda yakin?", vbYesNo, "Keluar")
    If a = vbYes Then
        edit_stok4
        F01_Main.Enabled = True
        F01_Main.Show
    Else
        Cancel = True
    End If
Else
    F01_Main.Enabled = True
    F01_Main.Show
End If
End Sub

Private Sub ListView1_Click()
isi_dtl
End Sub

Sub isi_dtl()
Dim kode As String
Dim nama As String
Dim cek As Boolean
Dim stat As String
cek = False
If Not dt_list.Recordset.BOF Then
    nama = ListView1.SelectedItem
    stat = ListView1.SelectedItem.SubItems(1)
    dt_produk.Refresh
    With dt_produk.Recordset
        If Not .BOF Then
            .MoveFirst
            Do Until cek = True
                If nama = !nama_produk Then
                    kode = !kode_produk
                    cek = True
                    .MovePrevious
                End If
                .MoveNext
            Loop
        End If
    End With
    dt_list.Refresh
    With dt_list.Recordset
    If Not .BOF Then
        cek = False
        .MoveFirst
        Do Until cek = True
            If kode = !kode_produk And stat = !status_beli Then
                cek = True
                .MovePrevious
            End If
            .MoveNext
        Loop
        Combo3 = nama
        Text15 = !qty_jual
        Text16 = !harga_satuan
        If !status_beli = "ISI ULANG" Then
            Option2(0).Value = True
        Else
            Option2(1).Value = True
            If !status_beli = "TABUNG+ISI" Then
                Option2(3).Value = True
            Else
                Option2(2).Value = True
            End If
        End If
    End If
    End With
End If
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Combo2.Visible = True
    Text6.Visible = True
    Label1(6).Visible = True
    Label1(7).Visible = True
    Label1(11).Visible = True
    Text8(1).Visible = True
    isi_lokasi
    hit_ttl
Case 1
    Combo2.Visible = False
    Text6.Visible = False
    Label1(6).Visible = False
    Label1(7).Visible = False
    Label1(11).Visible = False
    Text8(1).Visible = False
    Text8(1) = 0
    isi_lokasi
    hit_ttl
End Select
End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
Case 0
    Frame1(1).Visible = False
Case 1
    Frame1(1).Visible = True
Case 2
Case 3

End Select
'hit_trans
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text15_Change()
'hit_trans
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text7_Change()
isi_harga
isi_stok
End Sub

Private Sub Text8_Change(Index As Integer)
Select Case Index
Case 0, 1
'    hit_ttl
End Select
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
If Command1(0).Caption = "SIMPAN" Then
    If Me.Height < 8910 Then
        Me.Height = Me.Height + 267
    Else
        Timer1.Enabled = False
    End If
Else
    If Me.Height > 6240 Then
        Me.Height = Me.Height - 267
    Else
        Timer1.Enabled = False
    End If
End If
End Sub

Sub cmd_awal()
Command1(0).Picture = ImageList1.ListImages(1).Picture
Command1(1).Picture = ImageList1.ListImages(3).Picture
Command1(0).Caption = "TAMBAH"
Command1(1).Caption = "EDIT"
Command1(2).Visible = True
Command1(3).Visible = True
Command1(4).Visible = True
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
ListView1.Enabled = False
End Sub

Sub kosong1()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text8(0) = 0
Text8(1) = 0
Text8(2) = 0
Combo1.Clear
Combo2.Clear
ListView1.ListItems.Clear
End Sub

Sub kosong2()
Text7 = ""
Text9 = ""
Text10 = ""
Text11 = ""
Text12 = ""
Text13 = ""
Text14 = ""
Text15 = ""
Text16 = ""
Combo3.Clear
End Sub

Sub tutup1()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
End Sub

Sub tutup2()
Text7.Enabled = False
'Text9.Enabled = False
'Text10.Enabled = False
'Text11.Enabled = False
'Text12.Enabled = False
'Text13.Enabled = False
'Text14.Enabled = False
'Text15.Enabled = False
'Text16.Enabled = False
End Sub

Sub isi_cmb1()
Combo1.Clear
dt_member.RecordSource = "select * from member " 'order by nama_member"
dt_member.Refresh
With dt_member.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo1.AddItem !no_member
            .MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
End With
End Sub

Sub isi_member()
Dim cari As Boolean
cari = False
With dt_member.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Combo1 = !no_member Then
            Text1 = !nama_member
            Text2 = !alamat_member
            Text3 = !telp_member
            Text4 = !kode_lokasi
            kd_lok = !kode_lokasi
            cari = True
            .MoveLast
        End If
        .MoveNext
    Loop
    If cari = False Then
        Text1 = ""
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
    Else
        isi_lokasi
    End If
End If
End With
End Sub

Sub isi_lokasi()
dt_lokasi.Refresh
If Command3.Visible = True Then
    kd_lok = Text4
End If
With dt_lokasi.Recordset
If Not .BOF Then
dt_lokasi.Recordset.MoveFirst
If kd_lok <> "" Then
Do Until dt_lokasi.Recordset!kode_lokasi = kd_lok
    dt_lokasi.Recordset.MoveNext
Loop
End If
Text5 = dt_lokasi.Recordset!nama_lokasi
If Option1(0).Value = True Then
'    Text8(1) = Format(dt_lokasi.Recordset!biaya_kirim, "###,###.00")
    Text8(1) = 0
Else
    Text8(1) = 0
End If
End If
End With
End Sub

Sub hit_ttl()
Dim sub_ttl As Double
Dim biaya As Double
Text8(0) = Format(Text8(0), "###")
Text8(1) = Format(Text8(1), "###")
sub_ttl = Val(Text8(0))
biaya = Val(Text8(1))
Text8(2) = Format(sub_ttl + biaya, "###,###.00")
Text8(0) = Format(Text8(0), "###,###.00")
Text8(1) = Format(Text8(1), "###,###.00")
End Sub

Sub isi_cmb2()
Combo2.Clear
dt_petugas.RecordSource = "select * from petugas order by nama_petugas"
dt_petugas.Refresh
With dt_petugas.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo2.AddItem !nama_petugas
            .MoveNext
        Loop
        Combo2.ListIndex = 0
    End If
End With
End Sub

Sub isi_petugas()
dt_petugas.Refresh
With dt_petugas.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until !nama_petugas = Combo2
        .MoveNext
    Loop
    Text6 = !kode_petugas
End If
End With
End Sub

Sub isi_cmb3()
Combo3.Clear
dt_produk.RecordSource = "select * from produk"
dt_produk.Refresh
With dt_produk.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo3.AddItem !nama_produk
            .MoveNext
        Loop
        Combo3.ListIndex = 0
    End If
End With
End Sub

Sub isi_produk()
dt_produk.Refresh
With dt_produk.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until !nama_produk = Combo3
        .MoveNext
    Loop
    Text7 = !kode_produk
End If
End With
End Sub

Sub isi_harga()
dt_harga.Refresh
With dt_harga.Recordset
If Not Text7 = "" And Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Text7 = !kode_produk Then
            Text13 = Format(!harga_kosong, "###,###.00")
            Text14 = Format(!harga_isi, "###,###.00")
            Text12 = Format(!harga_kosong + !harga_isi, "###,###.00")
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub isi_stok()
Dim ttl As Single
Dim isi As Single
Dim kosong As Single
If Not Text7 = "" Then
dt_stok.RecordSource = "select * from stok where kode_produk='" & Text7 & "' order by tanggal_masuk,jam asc"
dt_stok.Refresh
With dt_stok.Recordset
    ttl = 0
    isi = 0
    kosong = 0
    If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        ttl = ttl + !jumlah
        isi = isi + !qty_isi
        kosong = kosong + !qty_kosong
        .MoveNext
    Loop
    End If
End With
Text11 = isi + kosong
Text9 = isi
Text10 = kosong
End If
End Sub

Sub hit_trans()
Dim jml As Double
jml = 0
If Text15 <> "" Then
If Option2(0).Value = True Then
    jml = Val(Text15) * Format(Text14, "###")
Else
    If Option2(3).Value = True Then
        jml = Val(Text15) * Format(Text12, "###")
    Else
        jml = Val(Text15) * Format(Text13, "###")
    End If
End If
End If
Text16 = Format(jml, "###,###.00")
End Sub

Sub cek_stok()
cek = True
If Text15 <> "" Then
If tambah = True Then
    If Option2(0).Value = True Then
        If Val(Text15) > Val(Text9) Then
            cek = False
        End If
    Else
        If Option2(3).Value = True Then
            If Val(Text15) > Val(Text9) Then
                cek = False
            End If
        Else
            If Val(Text15) > Val(Text10) Then
                cek = False
            End If
        End If
    End If
Else
    If Option2(0).Value = True Then
        If Val(Text15) - dt_list.Recordset!qty_jual > Val(Text9) Then
            cek = False
        End If
    Else
        If Option2(3).Value = True Then
            If Val(Text15) - dt_list.Recordset!qty_jual > Val(Text9) Then
                cek = False
            End If
        Else
            If Val(Text15) - dt_list.Recordset!qty_jual > Val(Text10) Then
                cek = False
            End If
        End If
    End If
End If
End If
End Sub

Private Sub Timer2_Timer()
Label2(3).Caption = Format(Date, "d mmmm yyyy") & ", " & Format(Time, "hh:mm:ss")
End Sub

Sub ISI_CMB4()
Combo4.Clear
dt_lokasi.Refresh
With dt_lokasi.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo4.AddItem !kode_lokasi
            .MoveNext
        Loop
        Combo4.ListIndex = 0
    End If
End With
End Sub

