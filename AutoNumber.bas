Attribute VB_Name = "AutoNumber"
Public urut_remain As String

Public Function produk_auto()
Dim urutan As String * 10
Dim hitung As Single
With F03_Produk.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "PRD" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("KOde_produk"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("kode_produk"), 7)) + 1
            urutan = "PRD" & Right("0000000" & hitung, 7)
        End If
    End If
    F03_Produk.Text1 = urutan
End With
End Function

Public Function supplier_auto()
Dim urutan As String * 10
Dim hitung As Single
With F04_Supplier.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "SPL" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("KOde_supplier"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("kode_supplier"), 7)) + 1
            urutan = "SPL" & Right("0000000" & hitung, 7)
        End If
    End If
    F04_Supplier.Text1(0) = urutan
End With
End Function

Public Function member_auto()
Dim urutan As String * 10
Dim hitung As Single
With F06_Member.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "AWP" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("no_member"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("no_member"), 7)) + 1
            urutan = "AWP" & Right("0000000" & hitung, 7)
        End If
    End If
    F06_Member.Text1(0) = urutan
End With
End Function

Public Function petugas_auto()
Dim urutan As String * 10
Dim hitung As Single
With F05_Petugas.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "PGS" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("KOde_petugas"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("kode_petugas"), 7)) + 1
            urutan = "PGS" & Right("0000000" & hitung, 7)
        End If
    End If
    F05_Petugas.Text1(0) = urutan
End With
End Function

Public Function Rkanan(ndata, cformat) As String
    Rkanan = Format(ndata, cformat)
    Rkanan = Space(Len(cformat) - Len(Rkanan)) + Rkanan
End Function

Public Function nota_auto()
Dim urutan As String * 12
Dim hitung As Single
With F09_Nota.dt_nota.Recordset
    If .RecordCount = 0 Then
        urutan = "Nota." & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("no_nota"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("no_nota"), 7)) + 1
            urutan = "Nota." & Right("0000000" & hitung, 7)
        End If
    End If
    F09_Nota.Caption = urutan
End With
End Function

Public Function remain_auto()
Dim urutan As String * 10
Dim hitung As Single
With F15_Remain.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "TAG" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("nomor"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("nomor"), 7)) + 1
            urutan = "TAG" & Right("0000000" & hitung, 7)
        End If
    End If
    urut_remain = urutan
End With
End Function

