Attribute VB_Name = "bytestokb"
Public Enum BYTEVALUES
  KiloByte = 1024
  MegaByte = 1048576
  GigaByte = 107374182
End Enum

Function ConvertBytes(Bytes As Long, Optional NumDigitsAfterDecmal As Long = 0) As String
  ConvertBytes = FormatNumber(Bytes, NumDigitsAfterDecmal) & "bytes"
End Function
Function ConvertKiloBytes(Bytes As Long, Optional NumDigitsAfterDecmal As Long = 0) As String
  ConvertKiloBytes = FormatNumber(Bytes / BYTEVALUES.KiloByte, NumDigitsAfterDecmal) & "kb"
End Function
Function ConvertMegaBytes(Bytes As Long, Optional NumDigitsAfterDecmal As Long = 0) As String
  ConvertMegaBytes = FormatNumber(Bytes / BYTEVALUES.MegaByte, NumDigitsAfterDecmal) & "mb"
End Function
Function ConvertGigaBytes(Bytes As Long, Optional NumDigitsAfterDecmal As Long = 0) As String
  ConvertGigaBytes = FormatNumber(Bytes / BYTEVALUES.GigaByte, NumDigitsAfterDecmal) & "gb"
End Function
