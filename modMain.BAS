Attribute VB_Name = "modMain"
Option Explicit
DefObj A-Z

'====================================================================
Public Sub Main()
  Dim Buff() As Byte
  Dim TXT1 As String
  Dim TXT2 As String
  Dim CL As New CCryptLibSupprt
  
  Open "~$anyfile.dat" For Binary Access Read Lock Write As #1
    ReDim Buff(0 To LOF(1) - 1)
    Get #1, , Buff
  Close #1
  
  TXT1 = StrConv(Buff, vbFromUnicode)
  
  CL.InitLibrary
  TXT2 = CL.EncryptData(TXT1, "Password", CRYPT_ALGO_3DES)
  CL.FreeLibrary
End Sub
