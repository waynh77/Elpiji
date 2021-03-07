VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form F11_TBayar 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Pembayaran"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   ClipControls    =   0   'False
   Icon            =   "F11_TBayar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   16325
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8421376
      ForeColor       =   12582912
      TabCaption(0)   =   "PEMBELIAN"
      TabPicture(0)   =   "F11_TBayar.frx":3482
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Data3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Data2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Data1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text10"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text8"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text6"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "DTPicker1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "ListView1(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "ComboBox2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "ComboBox1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Line1(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(14)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(13)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Line1(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(12)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label1(11)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label1(10)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label1(9)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label1(8)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label1(7)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label1(5)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Line1(1)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label1(4)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Line1(0)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label1(3)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label1(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label1(1)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Label1(0)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label1(6)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).ControlCount=   45
      TabCaption(1)   =   "PENJUALAN"
      TabPicture(1)   =   "F11_TBayar.frx":349E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(16)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ComboBox3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(17)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line1(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ComboBox4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ComboBox5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(18)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(19)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(20)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(21)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Line1(5)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line1(6)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1(22)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(23)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Line1(7)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1(24)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(25)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(26)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(27)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label1(28)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label1(29)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(30)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label1(15)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "DTPicker2"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "ListView1(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Command1(7)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Command1(8)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Command1(9)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Command1(10)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Command1(11)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Command1(12)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Command1(13)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Text12"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Text13"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Text14"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Text15"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Text16"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Text17"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Text18"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Text19"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Text20"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Text21"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Text22"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Data4"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Data5"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).ControlCount=   45
      Begin VB.Data Data5 
         Caption         =   "Data5"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   67
         Text            =   "Text3"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text21 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   66
         Text            =   "Text4"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text20 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8880
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   64
         Text            =   "Text6"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4800
         TabIndex        =   63
         Text            =   "Text7"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8880
         TabIndex        =   62
         Text            =   "Text8"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8880
         TabIndex        =   61
         Text            =   "Text9"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   2760
         TabIndex        =   60
         Text            =   "Text10"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   8160
         TabIndex        =   59
         Text            =   "Text11"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   2160
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   7560
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   1440
         Width           =   3255
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "F11_TBayar.frx":34BA
         Height          =   1695
         Left            =   -73200
         OleObjectBlob   =   "F11_TBayar.frx":34CE
         TabIndex        =   46
         Top             =   6720
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -73080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5520
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -73080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -73080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "KELUAR"
         DownPicture     =   "F11_TBayar.frx":3EA1
         Height          =   735
         Index           =   13
         Left            =   9960
         MouseIcon       =   "F11_TBayar.frx":460B
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":4915
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   8280
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "REFRESH"
         DownPicture     =   "F11_TBayar.frx":5997
         Height          =   735
         Index           =   12
         Left            =   9960
         MouseIcon       =   "F11_TBayar.frx":6101
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":640B
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   7560
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "CETAK"
         DownPicture     =   "F11_TBayar.frx":748D
         Height          =   735
         Index           =   11
         Left            =   9960
         MouseIcon       =   "F11_TBayar.frx":7BF7
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":7F01
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   6840
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "CARI"
         DownPicture     =   "F11_TBayar.frx":8F83
         Height          =   735
         Index           =   10
         Left            =   9960
         MouseIcon       =   "F11_TBayar.frx":96ED
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":99F7
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "HAPUS"
         DownPicture     =   "F11_TBayar.frx":AA79
         Height          =   735
         Index           =   9
         Left            =   9960
         MouseIcon       =   "F11_TBayar.frx":B1E3
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":B4ED
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5400
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "EDIT"
         DownPicture     =   "F11_TBayar.frx":C56F
         Height          =   735
         Index           =   8
         Left            =   9960
         MouseIcon       =   "F11_TBayar.frx":CCD9
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":CFE3
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "BAYAR"
         DownPicture     =   "F11_TBayar.frx":E065
         Height          =   735
         Index           =   7
         Left            =   9960
         MouseIcon       =   "F11_TBayar.frx":E7CF
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":EAD9
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "KELUAR"
         DownPicture     =   "F11_TBayar.frx":FB5B
         Height          =   735
         Index           =   6
         Left            =   -65040
         MouseIcon       =   "F11_TBayar.frx":102C5
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":105CF
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   8280
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "REFRESH"
         DownPicture     =   "F11_TBayar.frx":11651
         Height          =   735
         Index           =   5
         Left            =   -65040
         MouseIcon       =   "F11_TBayar.frx":11DBB
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":120C5
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   7560
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "CETAK"
         DownPicture     =   "F11_TBayar.frx":13147
         Height          =   735
         Index           =   4
         Left            =   -65040
         MouseIcon       =   "F11_TBayar.frx":138B1
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":13BBB
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6840
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "CARI"
         DownPicture     =   "F11_TBayar.frx":14C3D
         Height          =   735
         Index           =   3
         Left            =   -65040
         MouseIcon       =   "F11_TBayar.frx":153A7
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":156B1
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "HAPUS"
         DownPicture     =   "F11_TBayar.frx":16733
         Height          =   735
         Index           =   2
         Left            =   -65040
         MouseIcon       =   "F11_TBayar.frx":16E9D
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":171A7
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5400
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "EDIT"
         DownPicture     =   "F11_TBayar.frx":18229
         Height          =   735
         Index           =   1
         Left            =   -65040
         MouseIcon       =   "F11_TBayar.frx":18993
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":18C9D
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "BAYAR"
         DownPicture     =   "F11_TBayar.frx":19D1F
         Height          =   735
         Index           =   0
         Left            =   -65040
         MouseIcon       =   "F11_TBayar.frx":1A489
         MousePointer    =   99  'Custom
         Picture         =   "F11_TBayar.frx":1A793
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -66840
         TabIndex        =   27
         Text            =   "Text11"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -72240
         TabIndex        =   25
         Text            =   "Text10"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66120
         TabIndex        =   24
         Text            =   "Text9"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66120
         TabIndex        =   21
         Text            =   "Text8"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         TabIndex        =   19
         Text            =   "Text7"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   17
         Text            =   "Text6"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66120
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70200
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -67440
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -72840
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1440
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   -72840
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57212929
         CurrentDate     =   39710
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5055
         Index           =   0
         Left            =   -74880
         TabIndex        =   28
         Top             =   3960
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   8916
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5055
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   3960
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   8916
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   2160
         TabIndex        =   47
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57212929
         CurrentDate     =   39710
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T. BAYAR PENJUALAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   7680
         TabIndex        =   77
         Top             =   480
         Width           =   3105
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
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
         Height          =   315
         Index           =   30
         Left            =   120
         TabIndex        =   76
         Top             =   2040
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HARGA SATUAN"
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
         Height          =   315
         Index           =   29
         Left            =   2760
         TabIndex        =   75
         Top             =   2040
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "JUMLAH"
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
         Height          =   315
         Index           =   28
         Left            =   6840
         TabIndex        =   74
         Top             =   2040
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FREK"
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
         Height          =   315
         Index           =   27
         Left            =   120
         TabIndex        =   73
         Top             =   2400
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DUE DATE"
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
         Height          =   315
         Index           =   26
         Left            =   2760
         TabIndex        =   72
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "JUMLAH BAYAR"
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
         Height          =   315
         Index           =   25
         Left            =   6840
         TabIndex        =   71
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SISA/KURANG BAYAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   24
         Left            =   120
         TabIndex        =   70
         Top             =   3000
         Width           =   8745
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         Index           =   7
         X1              =   120
         X2              =   10800
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "STATUS PRODUK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   23
         Left            =   120
         TabIndex        =   69
         Top             =   3360
         Width           =   2625
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "STATUS PEMBAYARAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   22
         Left            =   5520
         TabIndex        =   68
         Top             =   3360
         Width           =   2625
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         Index           =   6
         X1              =   120
         X2              =   10800
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         Index           =   5
         X1              =   120
         X2              =   10800
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NO. MEMBER"
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
         Index           =   21
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NAMA PELANGGAN"
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
         Index           =   20
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Index           =   19
         Left            =   5520
         TabIndex        =   56
         Top             =   1080
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Index           =   18
         Left            =   5520
         TabIndex        =   55
         Top             =   1440
         Width           =   2025
      End
      Begin MSForms.ComboBox ComboBox5 
         Height          =   315
         Left            =   2160
         TabIndex        =   54
         Top             =   1080
         Width           =   1935
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3413;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox ComboBox4 
         Height          =   315
         Left            =   7560
         TabIndex        =   53
         Top             =   1080
         Width           =   1935
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3413;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         Index           =   4
         X1              =   120
         X2              =   10800
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NO. NOTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   315
         Index           =   17
         Left            =   3600
         TabIndex        =   50
         Top             =   480
         Width           =   2025
      End
      Begin MSForms.ComboBox ComboBox3 
         Height          =   315
         Left            =   5640
         TabIndex        =   49
         Top             =   480
         Width           =   1935
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3413;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TANGGAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   315
         Index           =   16
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   2025
      End
      Begin MSForms.ComboBox ComboBox2 
         Height          =   315
         Left            =   -67440
         TabIndex        =   45
         Top             =   1080
         Width           =   1935
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3413;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   315
         Left            =   -72840
         TabIndex        =   44
         Top             =   1080
         Width           =   1935
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3413;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         Index           =   3
         X1              =   -74880
         X2              =   -64200
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "STATUS PEMBAYARAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   14
         Left            =   -69480
         TabIndex        =   26
         Top             =   3360
         Width           =   2625
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "STATUS PRODUK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   13
         Left            =   -74880
         TabIndex        =   23
         Top             =   3360
         Width           =   2625
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         Index           =   2
         X1              =   -74880
         X2              =   -64200
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SISA/KURANG BAYAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   12
         Left            =   -74880
         TabIndex        =   22
         Top             =   3000
         Width           =   8745
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "JUMLAH BAYAR"
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
         Height          =   315
         Index           =   11
         Left            =   -68160
         TabIndex        =   20
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DUE DATE"
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
         Height          =   315
         Index           =   10
         Left            =   -72240
         TabIndex        =   18
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FREK"
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
         Height          =   315
         Index           =   9
         Left            =   -74880
         TabIndex        =   16
         Top             =   2400
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "JUMLAH"
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
         Height          =   315
         Index           =   8
         Left            =   -68160
         TabIndex        =   14
         Top             =   2040
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HARGA SATUAN"
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
         Height          =   315
         Index           =   7
         Left            =   -72240
         TabIndex        =   12
         Top             =   2040
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
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
         Height          =   315
         Index           =   5
         Left            =   -74880
         TabIndex        =   10
         Top             =   2040
         Width           =   1785
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         Index           =   1
         X1              =   -74880
         X2              =   -64200
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TRANSAKSI PEMBAYARAN STOK/PERSEDIAAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -71400
         TabIndex        =   9
         Top             =   480
         Width           =   7185
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         Index           =   0
         X1              =   -74880
         X2              =   -64200
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Index           =   3
         Left            =   -69480
         TabIndex        =   7
         Top             =   1440
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Index           =   2
         Left            =   -69480
         TabIndex        =   6
         Top             =   1080
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NAMA SUPPLIER"
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
         Index           =   1
         Left            =   -74880
         TabIndex        =   4
         Top             =   1440
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KODE SUPPLIER"
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
         Index           =   0
         Left            =   -74880
         TabIndex        =   3
         Top             =   1080
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TANGGAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   315
         Index           =   6
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   2025
      End
   End
End
Attribute VB_Name = "F11_TBayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub combobox1_Change()
isi_spl
End Sub

Sub isi_spl()
Data2.Refresh
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If ComboBox1 = !kode_supplier Then
            Text1 = !nama_supplier
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub isi_member()
Data5.Refresh
If ComboBox5 = "" And Not Data4.Recordset.BOF Then
    Text13 = ListView1(1).SelectedItem.ListSubItems(2).Text
Else
With Data5.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If ComboBox5 = !no_member Then
                Text13 = !nama_member
                .MoveLast
            End If
            .MoveNext
        Loop
    End If
End With
End If
End Sub

Sub isi_prd()
Data3.Refresh
With Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If ComboBox2 = !kode_produk Then
            Text2 = !nama_produk
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub isi_prd2()
Data3.Refresh
With Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If ComboBox4 = !kode_produk Then
            Text12 = !nama_produk
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Private Sub combobox1_Click()
isi_spl
End Sub

Private Sub combobox2_Change()
isi_prd
End Sub

Private Sub combobox2_Click()
isi_prd
End Sub

Private Sub ComboBox4_Change()
isi_prd2
End Sub

Private Sub ComboBox4_Click()
isi_prd2
End Sub

Private Sub ComboBox5_Change()
isi_member
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Text3 <> "" Then
        Me.Enabled = False
        F22_ByrBeli.Show
    Else
        MsgBox "Silahkan pilih data terlebih dahulu...", vbInformation, "Validasi Data"
    End If
Case 1
    Dim tgl As Date
    If Text3 <> "" Then
        Me.Enabled = False
        F23_EditByrBeli.Show
        With F23_EditByrBeli
            Call db_EDITbyrbeli
            .Data1.RecordSource = "select * from byrbeli where cdate(tgl_beli)='" & DTPicker1 & "' and kode_supplier='" & ComboBox1 & "' and kode_produk ='" & ComboBox2 & "' and qty = " & Text3
            .Data1.Refresh
            .kosong
            .isi_list
            .isi
        End With
    Else
        MsgBox "Silahkan pilih data terlebih dahulu...", vbInformation, "Validasi Data"
    End If
Case 2
'    Dim tgl2 As Date
    If Text3 <> "" Then
        Me.Enabled = False
        F23_EditByrBeli.Show
        With F23_EditByrBeli
            Call db_EDITbyrbeli
            .Data1.RecordSource = "select * from byrbeli where cdate(tgl_beli)='" & DTPicker1 & "' and kode_supplier='" & ComboBox1 & "' and kode_produk ='" & ComboBox2 & "' and qty = " & Text3
            .Data1.Refresh
            .kosong
            .isi_list
            .isi
        End With
    Else
        MsgBox "Silahkan pilih data terlebih dahulu...", vbInformation, "Validasi Data"
    End If
Case 3
    Call uncons
Case 4
    Call uncons
Case 5
    Call uncons
Case 6, 13
    Unload Me
Case 7
    If Text22 <> "" Then
        Me.Enabled = False
        F24_ByrJual.Show
    Else
        MsgBox "Silahkan pilih data terlebih dahulu...", vbInformation, "Validasi Data"
    End If
Case 8
    Dim tgl2 As Date
    If Text22 <> "" Then
        Me.Enabled = False
        F25_EditByrJual.Show
        With F25_EditByrJual
            Call db_EDITbyrJUAL
            .Data1.RecordSource = "select * from byrjual where cdate(tgl_jual)='" & DTPicker2 & "' and no_member='" & ComboBox5 & "' and kode_produk ='" & ComboBox4 & "' and qty = " & Text22
            .Data1.Refresh
            .kosong
            .isi_list
            .isi
        End With
    Else
        MsgBox "Silahkan pilih data terlebih dahulu...", vbInformation, "Validasi Data"
    End If
Case 9
    If Text22 <> "" Then
        Me.Enabled = False
        F25_EditByrJual.Show
        With F25_EditByrJual
            Call db_EDITbyrJUAL
            .Data1.RecordSource = "select * from byrjual where cdate(tgl_jual)='" & DTPicker2 & "' and no_member='" & ComboBox5 & "' and kode_produk ='" & ComboBox4 & "' and qty = " & Text22
            .Data1.Refresh
            .kosong
            .isi_list
            .isi
        End With
    Else
        MsgBox "Silahkan pilih data terlebih dahulu...", vbInformation, "Validasi Data"
    End If
Case 10
    Call uncons
Case 11
    Call uncons
Case 12
    Call uncons
End Select
End Sub

Private Sub Form_Activate()
isi_cmb1
isi_cmb2
'isi_cmb3
isi_cmb4
isi_cmb5
isi_list1
isi_list2
isi1
isi2
End Sub

Private Sub Form_Load()
Call db_TBayar
kosong1
tutup1
DTPicker1 = Date
kosong2
tutup2
DTPicker2 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
F01_Main.Enabled = True
F01_Main.Show
End Sub

Sub kosong1()
ComboBox1 = ""
ComboBox2 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text11 = ""
End Sub

Sub kosong2()
ComboBox3 = ""
ComboBox4 = ""
ComboBox5 = ""
Text13 = ""
Text12 = ""
Text22 = ""
Text21 = ""
Text20 = ""
Text19 = ""
Text18 = ""
Text17 = ""
Text16 = ""
Text15 = ""
Text14 = ""
End Sub

Sub tutup1()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
ComboBox1.Enabled = False
ComboBox2.Enabled = False
DTPicker1.Enabled = False
End Sub

Sub tutup2()
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Text19.Enabled = False
Text20.Enabled = False
Text21.Enabled = False
Text22.Enabled = False
ComboBox3.Enabled = False
ComboBox4.Enabled = False
ComboBox5.Enabled = False
DTPicker2.Enabled = False
End Sub

Sub buka1()
ComboBox1.Enabled = True
ComboBox2.Enabled = True
DTPicker1.Enabled = True
End Sub

Sub buka2()
ComboBox5.Enabled = True
ComboBox4.Enabled = True
DTPicker2.Enabled = True
End Sub

Sub isi1()
With Data1.Recordset
If Not .BOF Then
    DTPicker1 = !tgl_beli
    ComboBox1 = !kode_supplier
    ComboBox2 = !kode_produk
'    Text1 = !nama_supplier
'    Text2 = !nama_produk
    Text3 = !qty_beli
    Text4 = Format(!harga_satuan, "###,###.00")
    Text5 = Format(!qty_beli * !harga_satuan, "###,###.00")
    Text6 = !frek_bayar
    Text7 = Format(!due_date, "d mmmm yyyy")
    Text8 = Format(!jumlah_bayar, "###,###.00")
    Text9 = Format((!qty_beli * !harga_satuan) - !jumlah_bayar, "###,###.00")
    Text10 = !status_produk
    If !status_bayar = True Then
        Text11 = "LUNAS"
    Else
        Text11 = "BELUM LUNAS"
    End If
Else
    kosong1
End If
End With
End Sub

Sub isi2()
With Data4.Recordset
If Not .BOF Then
    DTPicker2 = !tgl_jual
    ComboBox5 = !no_member
    ComboBox4 = !kode_produk
    ComboBox3 = !no_nota
'    Text1 = !nama_supplier
'    Text2 = !nama_produk
    Text22 = !qty_jual
    Text21 = Format(!harga_satuan, "###,###.00")
    Text20 = Format(!qty_jual * !harga_satuan, "###,###.00")
    Text19 = !frek_bayar
    Text18 = Format(!due_date, "d mmmm yyyy")
    Text17 = Format(!jumlah_bayar, "###,###.00")
    Text16 = Format((!qty_jual * !harga_satuan) - !jumlah_bayar, "###,###.00")
    Text15 = !status_produk
    If !status_bayar = True Then
        Text14 = "LUNAS"
    Else
        Text14 = "BELUM LUNAS"
    End If
Else
    kosong2
End If
End With
End Sub

Sub isi_cmb1()
ComboBox1.Clear
Data2.Refresh
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        ComboBox1.AddItem !kode_supplier
        .MoveNext
    Loop
    ComboBox1.ListIndex = 0
End If
End With
End Sub

Sub isi_cmb5()
ComboBox5.Clear
Data5.Refresh
With Data5.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        ComboBox5.AddItem !no_member
        .MoveNext
    Loop
    ComboBox5.ListIndex = 0
End If
End With
End Sub

Sub isi_cmb2()
ComboBox2.Clear
Data3.Refresh
With Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        ComboBox2.AddItem !kode_produk
        .MoveNext
    Loop
    ComboBox2.ListIndex = 0
End If
End With
End Sub

Sub isi_cmb4()
ComboBox4.Clear
Data3.Refresh
With Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        ComboBox4.AddItem !kode_produk
        .MoveNext
    Loop
    ComboBox4.ListIndex = 0
End If
End With
End Sub

Sub isi_list1()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim ttl_qty As Single
Dim ttl_hrg As Double
ListView1(0).ColumnHeaders.Clear
ListView1(0).ListItems.Clear
Set head = ListView1(0).ColumnHeaders.Add(, , "Tanggal") ', ListView1(0).Width / 10 - 100)
Set head = ListView1(0).ColumnHeaders.Add(, , "Supplier", , 2) ' ListView1(0).Width / 10, 2)
Set head = ListView1(0).ColumnHeaders.Add(, , "Produk", , 2)  'ListView1(0).Width / 10, 2)
Set head = ListView1(0).ColumnHeaders.Add(, , "Status Produk", , 2)  ' ListView1(0).Width / 11, 2)
Set head = ListView1(0).ColumnHeaders.Add(, , "Status Bayar", , 2)  'ListView1(0).Width / 11, 2)
Set head = ListView1(0).ColumnHeaders.Add(, , "Frek", , 2)  'ListView1(0).Width / 11, 2)
Set head = ListView1(0).ColumnHeaders.Add(, , "Harga Satuan", , 1) ' ListView1(0).Width / 11, 1)
Set head = ListView1(0).ColumnHeaders.Add(, , "Qty", , 2)  'ListView1(0).Width / 11, 2)
Set head = ListView1(0).ColumnHeaders.Add(, , "Jumlah", , 1)  'ListView1(0).Width / 11, 1)
Set head = ListView1(0).ColumnHeaders.Add(, , "Jumlah Bayar", , 1) ' ListView1(0).Width / 11, 1)
Set head = ListView1(0).ColumnHeaders.Add(, , "Sisa Bayar", , 1)  'ListView1(0).Width / 11, 1)
ListView1(0).View = lvwReport
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1(0).ListItems.Add(, , Format(!tgl_beli, "d mmm yyyy"))
        dtl.SubItems(1) = !nama_supplier
        dtl.SubItems(2) = !nama_produk
        dtl.SubItems(3) = !status_produk
        If !status_bayar = True Then
            dtl.SubItems(4) = "LUNAS"
        Else
            dtl.SubItems(4) = "BELUM LUNAS"
        End If
        dtl.SubItems(5) = !frek_bayar
        dtl.SubItems(6) = Format(!harga_satuan, "###,###,00")
        dtl.SubItems(7) = !qty_beli
        dtl.SubItems(8) = Format(!harga_satuan * !qty_beli, "###,###,00")
        dtl.SubItems(9) = Format(!jumlah_bayar, "###,###,00")
        dtl.SubItems(10) = Format(!harga_satuan * !qty_beli - !jumlah_bayar, "###,###,00")
        .MoveNext
    Loop
End If
End With
Data1.Refresh
End Sub

Sub isi_list2()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim ttl_qty As Single
Dim ttl_hrg As Double
ListView1(1).ColumnHeaders.Clear
ListView1(1).ListItems.Clear
Set head = ListView1(1).ColumnHeaders.Add(, , "Tanggal") ', ListView1(1).Width / 10 - 100)
Set head = ListView1(1).ColumnHeaders.Add(, , "No Nota", , 2) ' ListView1(1).Width / 10, 2)
Set head = ListView1(1).ColumnHeaders.Add(, , "Member", , 2) ' ListView1(1).Width / 10, 2)
Set head = ListView1(1).ColumnHeaders.Add(, , "Produk", , 2)  'ListView1(1).Width / 10, 2)
Set head = ListView1(1).ColumnHeaders.Add(, , "Status Produk", , 2)  ' ListView1(1).Width / 11, 2)
Set head = ListView1(1).ColumnHeaders.Add(, , "Status Bayar", , 2)  'ListView1(1).Width / 11, 2)
Set head = ListView1(1).ColumnHeaders.Add(, , "Frek", , 2)  'ListView1(1).Width / 11, 2)
Set head = ListView1(1).ColumnHeaders.Add(, , "Harga Satuan", , 1) ' ListView1(1).Width / 11, 1)
Set head = ListView1(1).ColumnHeaders.Add(, , "Qty", , 2)  'ListView1(1).Width / 11, 2)
Set head = ListView1(1).ColumnHeaders.Add(, , "Jumlah", , 1)  'ListView1(1).Width / 11, 1)
Set head = ListView1(1).ColumnHeaders.Add(, , "Jumlah Bayar", , 1) ' ListView1(1).Width / 11, 1)
Set head = ListView1(1).ColumnHeaders.Add(, , "Sisa Bayar", , 1)  'ListView1(1).Width / 11, 1)
ListView1(1).View = lvwReport
With Data4.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1(1).ListItems.Add(, , Format(!tgl_jual, "d mmm yyyy"))
        dtl.SubItems(1) = !no_nota
        dtl.SubItems(2) = !nama_pembeli
        dtl.SubItems(3) = !nama_produk
        dtl.SubItems(4) = !status_produk
        If !status_bayar = True Then
            dtl.SubItems(5) = "LUNAS"
        Else
            dtl.SubItems(5) = "BELUM LUNAS"
        End If
        dtl.SubItems(6) = !frek_bayar
        dtl.SubItems(7) = Format(!harga_satuan, "###,###,00")
        dtl.SubItems(8) = !qty_jual
        dtl.SubItems(9) = Format(!harga_satuan * !qty_jual, "###,###,00")
        dtl.SubItems(10) = Format(!jumlah_bayar, "###,###,00")
        dtl.SubItems(11) = Format(!harga_satuan * !qty_jual - !jumlah_bayar, "###,###,00")
        .MoveNext
    Loop
End If
End With
Data4.Refresh
End Sub

Private Sub ListView1_Click(Index As Integer)
Select Case Index
Case 0
    gerak1
Case 1
    gerak2
End Select
End Sub

Sub gerak1()
Dim cek As Boolean
Dim tgl As Date
Dim spl As String
Dim prd As String
Dim sts_prd As String
Dim qty As Single
tgl = ListView1(0).SelectedItem.Text
spl = ListView1(0).SelectedItem.SubItems(1)
prd = ListView1(0).SelectedItem.SubItems(2)
sts_prd = ListView1(0).SelectedItem.SubItems(3)
qty = ListView1(0).SelectedItem.SubItems(7)
cek = False
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until cek = True Or .EOF
            If !tgl_beli = tgl And !nama_supplier = spl And !nama_produk = prd And !status_produk = sts_prd And !qty_beli = qty Then
                cek = True
                isi1
                .MovePrevious
            End If
            .MoveNext
        Loop
    End If
End With
End Sub

Sub gerak2()
Dim cek As Boolean
Dim tgl As Date
Dim mbr As String
Dim prd As String
Dim sts_prd As String
Dim qty As Single
If Not Data4.Recordset.BOF Then
tgl = ListView1(1).SelectedItem.Text
mbr = ListView1(1).SelectedItem.SubItems(2)
prd = ListView1(1).SelectedItem.SubItems(3)
sts_prd = ListView1(1).SelectedItem.SubItems(4)
qty = ListView1(1).SelectedItem.SubItems(8)
Text13 = mbr
cek = False
With Data4.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until cek = True Or .EOF
            If !tgl_jual = tgl And !nama_pembeli = mbr And !nama_produk = prd And !status_produk = sts_prd And !qty_jual = qty Then
                cek = True
                isi2
                .MovePrevious
            End If
            .MoveNext
        Loop
    End If
End With
End If
End Sub

Private Sub ListView1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
Case 0
    gerak1
Case 1
    gerak2
End Select
End Sub

Private Sub ListView1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
Case 0
    gerak1
Case 1
    gerak2
End Select
End Sub
