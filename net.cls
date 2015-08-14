VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub ShowSelected()
    Dim netwerk As Double

    netwerk = IPToNumber(ColVal("netwerk"))
    subnet = NetSize(CInt(Split(ColVal("sub"), "/")(1)))
    gateway = IPToNumber(ColVal("gateway"))
    Debug.Print netwerk, subnet, gateway
    If XMod(netwerk, subnet) = 0 Then
        Debug.Print "Netwerkadres OK"
    Else
        Debug.Print "Netwerkadres FOUT"
    End If
    If gateway > netwerk + 1 And gateway < netwerk + subnet - 1 Then
        Debug.Print "gateway ok"
    Else
        Debug.Print "gateway error"
    End If

End Sub


Function CelVal(kol As Integer, rij As Integer) As String
    GetCell = ActiveSheet.Cells.Item(kol, rij)
    Return
End Function

Function ColVal(colname As String) As String
    For kol = 1 To 30
        cel = ActiveSheet.Cells.Item(1, kol)
        If cel = colname Then Exit For
    Next
    
    If kol < 30 Then
        ColVal = ActiveSheet.Cells.Item(ActiveCell.Row, kol)
    Else
        ColVal = ""
    End If
End Function

Function IPToNumber(ip As String) As Double
    ' test eerst of het wel een IP adres is
    x = Split(ip, ".")
    IPToNumber = 0
    For i = 0 To UBound(x)
        IPToNumber = IPToNumber * 256
        IPToNumber = IPToNumber + x(i)
    Next i
End Function

Function NetSize(subnet As Integer) As Integer
    NetSize = 1
    For i = 31 To subnet Step -1
        NetSize = NetSize * 2
    Next
End Function

Function NumberToIP(ip As Double) As String
    For i = 1 To 4
        NumberToIP = NumberToIP & "." & XMod(ip, 256)
        ip = ip / 256
    Next
    NumberToIP = Right(NumberToIP, Len(NumberToIP) - 1)
End Function

Function XMod(ByVal Number As Double, ByVal Divisor As Double) As Double
    Number = Int(Abs(Number))
    Divisor = Int(Abs(Divisor))
    XMod = Number - (Int(Number / Divisor) * Divisor)
End Function
