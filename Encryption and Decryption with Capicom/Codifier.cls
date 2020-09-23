VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Codifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'NOTE: i've segmented functions/subs as much as possible only to let the REMs be as much as possible readable,
'even if i HATE to "segment" code with underscore (_)

Public Function DecryptFileToText(ByVal eFileName As String, ByVal Key As String) As String
Dim tmpTxt As String
tmpTxt = RetrieveFileDataBodyText(eFileName)
DecryptFileToText = DecryptText(tmpTxt, Key)
End Function

Public Sub EncryptTextToFile(ByVal eFileName As String, ByVal dText As String, ByVal Key As String)
Dim tmpTxt As String

tmpTxt = EncryptText(dText, Key)
WriteToFile eFileName, tmpTxt
End Sub


Public Function DecryptFile(ByVal dFileName As String, _
   ByVal decFileName, _
   ByVal Key As String) As Long
   
DecFile dFileName, decFileName, Key                                     'Just look forward...
                                                                        'this ensures the decrypting
                                                                        'routines !
End Function

Public Function EncryptFile(ByVal eFileName As String, _
   ByVal dFileName, _
   ByVal Key As String) As Long
   
EncryptFile = EncFile(eFileName, dFileName, Key)                        'Just look forward...
                                                                        'this ensures the encrypting
                                                                        'routines !
End Function

Public Function DecryptText(ByVal dText As String, ByVal Key As String) As String
Dim tmp As String                                                       'set a dummy variable

tmp = TextDecrypt(dText, Key)                                           'Decrypt data using a temp variable
DecryptText = IIf(tmp = vbNullString, "(Wrong password)", tmp)          'Decide the results

End Function

Public Function EncryptText(ByVal eText As String, ByVal Key As String) As String
EncryptText = TextEncrypt(eText, Key)                                   'Go to encrypt text
End Function

Private Function TextEncrypt(ByVal eText As String, ByVal Key As String) As String
Dim objCrypto As New EncryptedData                                      'instance to CAPICOM obj
   
   If eText <> vbNullString And Key <> vbNullString Then                'if non-valid data, exits
      With objCrypto
         .Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_3DES            'select algorithm
         .Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS  'select key length
         .SetSecret Key                                                 'set the provided password
         .Content = eText                                               'set data to encrypt
         TextEncrypt = .Encrypt(CAPICOM_ENCODE_BASE64)                  'encrypt !
      End With
    End If

End Function

Private Function TextDecrypt(ByVal dText As String, ByVal Key As String) As String
Dim objCrypto As New EncryptedData                                      'instance to CAPICOM obj
   On Error Resume Next                                                 'to prevent non-valid decrypting results
   
   If dText <> "" And Key <> "" Then                                    'if non-valid data, exits
      With objCrypto
         .SetSecret Key                                                 'set the provided password
         .Decrypt dText                                                 'decrypt the provided scrmbled text !
         TextDecrypt = .Content                                         'save the unscrmbled text
      End With
   End If

End Function

Private Function EncFile(ByVal srcFileName As String, ByVal destFileName As String, ByVal encKey As String) As Long
Dim FileSS As String, tmpString As String
   
   If srcFileName <> destFileName Then                                  'simple an easy way to prevent
                                                                        'cases in which source and
                                                                        'dest are the same file ! btw a very stupid but extremely fast way 8-)
      'Retrieve data
      FileSS = RetrieveFileDataBodyText(srcFileName)                    'get the entire file using api
      
      tmpString = TextEncrypt(FileSS, encKey)                           'and Encrypt it
      
      WriteToFile destFileName, tmpString                               'create new file which contains
                                                                        'the brand-new encrypted data !
   End If
   
End Function

Private Sub DecFile(ByVal srcFileName As String, ByVal destFileName As String, ByVal encKey As String)
Dim ValRit As Long, bWritten As Long, objHandle As Long
Dim FileSS As String, tmpString As String

   If srcFileName <> destFileName Then
      'Retrieve data
      FileSS = RetrieveFileDataBodyText(srcFileName)
      
      '2) encrypt
      tmpString = TextDecrypt(FileSS, encKey)                           'The same of EncFile: the
                                                                        'only difference is that
                                                                        'now we decrypt the file !
      
      '3) create new encrypted file
      WriteToFile destFileName, tmpString
   End If

End Sub


Private Function RetrieveFileDataBodyText(ByVal tgtFileName As String) As String
Dim lngFileHandle As Long, lngSize As Long, lngRet As Long
Dim ofData As OFSTRUCT, FileSS As String

   lngFileHandle = OpenFile(tgtFileName, ofData, 0)                     'Creates an handle to the file provided
   lngSize = GetFileSize(lngFileHandle, 0)                              'retrieve it's size (if 0 or non exists it fail)
   
   If lngSize > 0 Then                                                  'ok, keep on working
      ReDim arrFileMain(lngSize) As Byte                                'it will contain streamed data from file
      
                                                                        'Now i binary retrieve data from file
      ReadFile lngFileHandle, _
               arrFileMain(0), _
               UBound(arrFileMain), _
               lngRet, _
               ByVal 0&
      
      FileSS = Space$(UBound(arrFileMain))                              'Prepare a string with blanks to store (just red ! ) data
      
                                                                        'Inflating the above variable with real data
      CopyMemory ByVal StrPtr(FileSS), _
            arrFileMain(0), _
            UBound(arrFileMain)
                                                                        'Exit the function with required data to perform
                                                                        'next steps...
      RetrieveFileDataBodyText = _
         Left$(StrConv(FileSS, vbUnicode), _
         lngSize)
      
      CloseHandle lngFileHandle                                         'Close the opened handle to the file above
   End If

End Function

Private Function WriteToFile(ByVal destFile As String, ByVal Text As String) As Long
Dim ValRit As Long, bWritten As Long, objHandle As Long
Dim ofData As OFSTRUCT

   'Write directly data to file in one time
   objHandle = CreateFile(destFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, 0, 0)
   ValRit = WriteFile(objHandle, ByVal Text, Len(Text), bWritten, ByVal 0&)
                        
   'closes handle and terminate write process
   CloseHandle objHandle


End Function
