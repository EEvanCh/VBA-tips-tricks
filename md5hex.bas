Function StringToMD5Hex(ByVal s As String) As String
Dim enc As Object
Dim bytes() As Byte
Dim pos As Long
Dim outstr As String

Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

bytes = StrConv(s, vbFromUnicode)
bytes = enc.ComputeHash_2(bytes)

For pos = LBound(bytes) To UBound(bytes)
   outstr = outstr & LCase(Right("0" & Hex(bytes(pos)), 2))
Next pos

StringToMD5Hex = outstr
Set enc = Nothing
End Function
