VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCryptLibSupprt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Initialized As Boolean

Private Type LIBRARY_INFO
  Description As String
  MajorVersion As Long
  MinorVersion As Long
  RevisionVersion As Long
  Copyright As String
End Type

Private Const PUSH_BUFFER_SIZE As Long = 2& * 1024& ' 2Kb
Private Const POP_BUFFER_SIZE As Long = 2& * 1024&  ' 2Kb

'====================================================================
Public Sub InitLibrary()
  Dim RC As Long
  
  If Initialized Then Exit Sub
  
  RC = cryptInit()
  If RC <> CRYPT_OK Then
    Err.Raise vbObjectError, , _
              "Error initilizing CryptLib library." & _
              GetCLErrorMessage(RC)
  Else
    Initialized = True
  End If
End Sub

Public Sub FreeLibrary()
  Dim RC As Long
  
  If Not Initialized Then Exit Sub
  
  RC = cryptEnd()
  If RC <> CRYPT_OK Then
    Err.Raise vbObjectError, , _
              "Error releasing CryptLib library." & _
              GetCLErrorMessage(RC)
  Else
    Initialized = False
  End If
End Sub

'====================================================================
Public Function EncryptData(ByVal Text As String, _
                            ByVal password As String, _
                            Optional ByVal Algorithm As CRYPT_ALGO_TYPE = CRYPT_ALGO_AES) As String
  Debug.Print "==> EncryptData()"
  EncryptData = EncryptDecryptData(Text, password, True, Algorithm)
End Function

Public Function DecryptData(ByVal Text As String, _
                            ByVal password As String) As String
  Debug.Print "==> DecryptData()"
  DecryptData = EncryptDecryptData(Text, password, False)
End Function

'====================================================================
Private Function EncryptDecryptData(ByVal Text As String, _
                                    ByVal password As String, _
                                    ByVal EncryptIt As Boolean, _
                                    Optional ByVal Algorithm As CRYPT_ALGO_TYPE = CRYPT_ALGO_AES) As String
  Dim hEnvelope As Long
  Dim RC As Long
  Dim ErrText As String
  Dim CharIndex As Long
  Dim ResText As String
  Dim BytesPushed As Long
  
  Debug.Print "==> EncryptDecryptData()"
  
  If Not Initialized Then
    ErrText = "CryptLib library was not initialized"
    GoTo hFinally
  End If
  
  RC = cryptCreateEnvelope(hEnvelope, CRYPT_UNUSED, _
                           IIf(EncryptIt, CRYPT_FORMAT_CRYPTLIB, CRYPT_FORMAT_AUTO))
  If RC < CRYPT_OK Then
    ErrText = "Error creating Envelope object." & GetCLErrorMessage(RC)
    GoTo hFinally
  End If
  
  If EncryptIt Then
    RC = cryptSetAttribute(hEnvelope, CRYPT_OPTION_ENCR_ALGO, Algorithm)
    If RC < CRYPT_OK Then
      ErrText = "Can't set crypt algorithm." & GetCLErrorMessage(RC)
      GoTo hFinally
    End If
  End If
  
  If EncryptIt And password <> "" Then
    RC = cryptSetAttributeString(hEnvelope, CRYPT_ENVINFO_PASSWORD, password, Len(password))
    If RC < CRYPT_OK Then
      ErrText = "Can't set password." & GetCLErrorMessage(RC)
      GoTo hFinally
    End If
  End If
  
  If EncryptIt Then
    RC = cryptSetAttribute(hEnvelope, CRYPT_ENVINFO_DATASIZE, Len(Text))
    If RC < CRYPT_OK Then
      ErrText = "Can't set data size." & GetCLErrorMessage(RC)
      GoTo hFinally
    End If
  End If
  
  CharIndex = 1
  Do While CharIndex < Len(Text)
    RC = PushData(hEnvelope, Mid(Text, CharIndex, PUSH_BUFFER_SIZE), password, ErrText, BytesPushed)
    If RC < CRYPT_OK And RC <> CRYPT_ERROR_OVERFLOW Then GoTo hFinally
    
    If RC = CRYPT_ERROR_OVERFLOW Or BytesPushed = 0 Then
      RC = PopData(hEnvelope, ResText, ErrText)
      If RC < CRYPT_OK Then GoTo hFinally
      
      EncryptDecryptData = EncryptDecryptData & ResText
    End If
    
    CharIndex = CharIndex + BytesPushed
  Loop
  
  ResText = ""
  Do
    RC = FlushData(hEnvelope, ErrText)
    If RC < CRYPT_OK Then GoTo hFinally
    
    RC = PopData(hEnvelope, ResText, ErrText)
    If RC < CRYPT_OK Then GoTo hFinally
    
    EncryptDecryptData = EncryptDecryptData & ResText
  Loop While ResText <> ""
  
  RC = CRYPT_OK
  ErrText = ""
  
hFinally:
  If hEnvelope > 0 Then _
    RC = cryptDestroyEnvelope(hEnvelope)
  
  If ErrText <> "" Then Err.Raise vbObjectError, , ErrText
End Function

'====================================================================
Private Function PushData(ByVal hEnvelope As Long, _
                          ByVal Text As String, _
                          ByVal password As String, _
                          ByRef lpErrText As String, _
                          ByRef lpBytesCopied As Long) As Long
  Dim RC As Long
  
  RC = cryptPushData(hEnvelope, Text, Len(Text), lpBytesCopied)
  Debug.Print "cryptPushData() ==> len: " & Len(Text) & ", copied: " & lpBytesCopied
  If RC < CRYPT_OK Then
    If RC = CRYPT_ENVELOPE_RESOURCE Then
      If password = "" Then
        lpErrText = "Can't ""push"" data - password required." & GetCLErrorMessage(RC)
        GoTo hFinally
      Else
        RC = cryptSetAttributeString(hEnvelope, CRYPT_ENVINFO_PASSWORD, password, Len(password))
        If RC < CRYPT_OK Then
          lpErrText = "Can't set password. Is it valid?" & GetCLErrorMessage(RC)
          GoTo hFinally
        End If
        RC = CRYPT_OK
      End If
    Else
      lpErrText = "Can't ""push"" data." & GetCLErrorMessage(RC)
      GoTo hFinally
    End If
  End If
  
  RC = CRYPT_OK
hFinally:
  PushData = RC
End Function

'====================================================================
Private Function FlushData(ByVal hEnvelope As Long, _
                           ByRef lpErrText As String) As Long
  Dim bytesCopied As Long
  Dim RC As Long
  
  RC = cryptFlushData(hEnvelope)
  Debug.Print "cryptFlushData()"
  If RC < CRYPT_OK Then
    lpErrText = "Can't flush internal data." & GetCLErrorMessage(RC)
    GoTo hFinally
  End If
  
  RC = CRYPT_OK
hFinally:
  FlushData = RC
End Function

'====================================================================
Private Function PopData(ByVal hEnvelope As Long, _
                         ByRef lpText As String, _
                         ByRef lpErrText As String) As Long
  Dim bytesCopied As Long
  Dim EnvelopedData As String
  Dim nEnvelopedData As Long
  Dim RC As Long
  
  nEnvelopedData = POP_BUFFER_SIZE
  EnvelopedData = String(nEnvelopedData + 1, 0)
  
  bytesCopied = 0
  lpText = ""
  
  Do
    RC = cryptPopData(hEnvelope, EnvelopedData, nEnvelopedData, bytesCopied)
    Debug.Print "cryptPopData() ==> requested: " & nEnvelopedData & ", copied: " & bytesCopied
    If RC < CRYPT_OK Then
      lpErrText = "Can't ""pop"" data." & GetCLErrorMessage(RC)
      GoTo hFinally
    End If
    
    lpText = lpText & Left(EnvelopedData, bytesCopied)
  Loop While bytesCopied > 0
  
  RC = CRYPT_OK
hFinally:
  PopData = RC
End Function

'====================================================================
Private Function GetCLErrorDescription(ByVal Number As Long) As String
  ' No error in function call
  Select Case Number
    Case CRYPT_OK
      GetCLErrorDescription = "No error"
      
    '  Error in parameters passed to function.  The parentheses are to catch
    '  potential erroneous use in an expression
    Case CRYPT_ERROR_PARAM1
      GetCLErrorDescription = "Bad argument, parameter 1"
    Case CRYPT_ERROR_PARAM2
      GetCLErrorDescription = "Bad argument, parameter 2"
    Case CRYPT_ERROR_PARAM3
      GetCLErrorDescription = "Bad argument, parameter 3"
    Case CRYPT_ERROR_PARAM4
      GetCLErrorDescription = "Bad argument, parameter 4"
    Case CRYPT_ERROR_PARAM5
      GetCLErrorDescription = "Bad argument, parameter 5"
    Case CRYPT_ERROR_PARAM6
      GetCLErrorDescription = "Bad argument, parameter 6"
    Case CRYPT_ERROR_PARAM7
      GetCLErrorDescription = "Bad argument, parameter 7"
      
    ' Errors due to insufficient resources
    Case CRYPT_ERROR_MEMORY
      GetCLErrorDescription = "Out of memory"
    Case CRYPT_ERROR_NOTINITED
      GetCLErrorDescription = "Data has not been initialised"
    Case CRYPT_ERROR_INITED
      GetCLErrorDescription = "Data has already been initialised"
    Case CRYPT_ERROR_NOSECURE
      GetCLErrorDescription = "Opnion not available at requested security level"
    Case CRYPT_ERROR_RANDOM
      GetCLErrorDescription = "No reliable random data available"
    Case CRYPT_ERROR_FAILED
      GetCLErrorDescription = "Operation failed"
    Case CRYPT_ERROR_INTERNAL
      GetCLErrorDescription = "Internal consistency check failed"
      
    ' Security violations
    Case CRYPT_ERROR_NOTAVAIL
      GetCLErrorDescription = "This type of opn.not available"
    Case CRYPT_ERROR_PERMISSION
      GetCLErrorDescription = "No permiss.to perform this operation"
    Case CRYPT_ERROR_WRONGKEY
      GetCLErrorDescription = "Incorrect key used to decrypt data"
    Case CRYPT_ERROR_INCOMPLETE
      GetCLErrorDescription = "Operation incomplete/still in progress"
    Case CRYPT_ERROR_COMPLETE
      GetCLErrorDescription = "Operation complete/can't continue"
    Case CRYPT_ERROR_TIMEOUT
      GetCLErrorDescription = "Operation timed out before completion"
    Case CRYPT_ERROR_INVALID
      GetCLErrorDescription = "Invalid/inconsistent information"
    Case CRYPT_ERROR_SIGNALLED
      GetCLErrorDescription = "Resource destroyed by extnl.event"
      
    ' High-level function errors
    Case CRYPT_ERROR_OVERFLOW
      GetCLErrorDescription = "Resources/space exhausted"
    Case CRYPT_ERROR_UNDERFLOW
      GetCLErrorDescription = "Not enough data available"
    Case CRYPT_ERROR_BADDATA
      GetCLErrorDescription = "Bad/unrecognised data format"
    Case CRYPT_ERROR_SIGNATURE
      GetCLErrorDescription = "Signature/integrity check failed"
      
    ' Data access function errors
    Case CRYPT_ERROR_OPEN
      GetCLErrorDescription = "Cannot open object"
    Case CRYPT_ERROR_READ
      GetCLErrorDescription = "Cannot read item from object"
    Case CRYPT_ERROR_WRITE
      GetCLErrorDescription = "Cannot write item to object"
    Case CRYPT_ERROR_NOTFOUND
      GetCLErrorDescription = "Requested item not found in object"
    Case CRYPT_ERROR_DUPLICATE
      GetCLErrorDescription = "Item already present in object"
      
    ' Data enveloping errors
    Case CRYPT_ENVELOPE_RESOURCE
      GetCLErrorDescription = "Need resource to proceed"
      
    Case Else
      GetCLErrorDescription = ""
  End Select
End Function

Private Function GetCLErrorMessage(ByVal Number As Long) As String
  Dim TXT As String
  
  If Number >= 0 Then
    GetCLErrorMessage = ""
    Exit Function
  End If
  
  TXT = GetCLErrorDescription(Number)
  If TXT = "" Then
    TXT = " Error code is " & CStr(Number) & "."
  Else
    TXT = " Error code is " & CStr(Number) & " (" & TXT & ")."
  End If
  
  GetCLErrorMessage = TXT
End Function
