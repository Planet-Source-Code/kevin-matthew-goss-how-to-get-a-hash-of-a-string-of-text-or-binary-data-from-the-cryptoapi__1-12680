<div align="center">

## How to get a hash of a string of text or binary data from the CryptoAPI


</div>

### Description

This code demonstrates how to get a hash, or a type of fingerprint for any string of data. Hashes are useful in determining if a piece of data has been altered. Usefull for files, data transmissions, etc.

Have fun with it!!!
 
### More Info
 
The function takes two parameters: (1) data - string of data to get hash of & (2) hashType - 0 or 1 - 0 is MD5 algorithm (16 character hash) and 1 is SHA (20 character hash)

You must have the proper version of IE installed (get 5 to be sure). The ADVAPI.dll that this uses does not function the same over all platforms, so expect to have to tweak this a bit.

returns ASCII string (could contain any ASCII characters)

Does not always print due to some ASCII characters not being printable. If you need to store it, use a binary file.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kevin Matthew Goss](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kevin-matthew-goss.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kevin-matthew-goss-how-to-get-a-hash-of-a-string-of-text-or-binary-data-from-the-cryptoapi__1-12680/archive/master.zip)

### API Declarations

```
'===============================================
' This code shows how to generate a hash of a
' string of data. There are 2 algorithms -'available in this verion - MD5 and SHA
' Hashes are extremely usefull for determining
'whether a transmission or file has been altered.'
' The MD5 returns a 16 character hash 'and the
'SHA returns a 20 character hash. No too hashes
' are alike unless the string matches perfectly,
' whether binary data or a text string.
' I use hashes to create crypto keys and to
'verify integrity of packets when using winsock
'(UDP especially). Be aware that hashes may not
'store to text correctly because of the possible
' existence of non printable characters in the
'stream - ues them runtime only or store them in
' a binary file - the APIs declarations are -'adapted from Davis Chapman -
'=================================================
' Algorithm classes
Private Const ALG_CLASS_ANY = 0
Private Const ALG_CLASS_SIGNATURE = 8192
Private Const ALG_CLASS_MSG_ENCRYPT = 16384
Private Const ALG_CLASS_DATA_ENCRYPT = 24576
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_CLASS_KEY_EXCHANGE = 40960
' Algorithm types
Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_DSS = 512
Private Const ALG_TYPE_RSA = 1024
Private Const ALG_TYPE_BLOCK = 1536
Private Const ALG_TYPE_STREAM = 2048
Private Const ALG_TYPE_DH = 2560
Private Const ALG_TYPE_SECURECHANNEL = 3072
' RC2 sub-ids
Private Const ALG_SID_RC2 = 2
' Stream cipher sub-ids
Private Const ALG_SID_RC4 = 1
Private Const ALG_SID_SEAL = 2
' Hash sub ids
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private cryptContext As Long
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL = 1
Private Const HP_HASHVAL = &H2
Private Const CRYPT_NEWKEYSET = &H8
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_SHA = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)
Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
 ByVal hProv As Long, _
 ByVal Algid As Long, _
 ByVal hKey As Long, _
 ByVal dwFlags As Long, _
 phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" ( _
 ByVal hHash As Long, _
 ByVal pbData As String, _
 ByVal dwDataLen As Long, _
 ByVal dwFlags As Long) As Long
'used to gain access to cryptoAPI
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
 phProv As Long, _
 ByVal pszContainer As String, _
 ByVal pszProvider As String, _
 ByVal dwProvType As Long, _
 ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
 ByVal hProv As Long, _
 ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" _
 (ByVal hProv As Long, _
 ByVal dwLen As Long, _
 ByVal pbBuffer As String) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
 ByVal hHash As Long, _
 ByVal dwParam As Long, _
 ByVal pbData As String, _
 pdwDataLen As Long, _
 ByVal dwFlags As Long) As Long
```


### Source Code

```
Public Function getHash(data As String, hashType As Integer) As String
Dim ht As Long
Dim sTemp As String
Dim sProv As String
Dim hLen As Long
Dim h As String
Dim hl As Long
'get hash type
If hashType = 0 Then
 'MD5
 ht = CALG_MD5
 hLen = 16
ElseIf hashType = 1 Then
 'SHA
 hLen = 20
 ht = CALG_SHA
Else
 getHash = ""
 Exit Function
End If
'--- Prepare string buffers
sTemp = vbNullChar
sProv = MS_DEF_PROV & vbNullChar
'---Gain Access To CryptoAPI
If Not CBool(CryptAcquireContext(cryptContext, sTemp, sProv, PROV_RSA_FULL, 0)) Then
 If Not CBool(CryptAcquireContext(cryptContext, sTemp, sProv, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
 getHash = ""
 Exit Function
 End If
End If
'Create Empty hash object
If Not CBool(CryptCreateHash(cryptContext, ht, 0, 0, hl)) Then
 getHash = ""
 Exit Function
End If
'Hash the input string.
If Not CBool(CryptHashData(hl, data, Len(data), 0)) Then
 getHash = ""
 Exit Function
End If
h = String(20, vbNull)
'Get hash val
If Not CBool(CryptGetHashParam(hl, HP_HASHVAL, h, hLen, 0)) Then
 getHash = ""
 Exit Function
End If
getHash = h
End Function
```

