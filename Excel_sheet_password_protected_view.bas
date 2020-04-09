Attribute VB_Name = "Module1"
Option Explicit
'+-------------------------------------------------------------------------------------------+
'¦   This is a VBA script that can be used to protect Excel Worksheets from both being viewed¦
'¦and edited with a password for each worksheet.                                             ¦
'¦                                                                                           ¦
'¦   For the sake of simplicity and easy user usage a shortcut key has to be defined for     ¦
'¦password_View_Sheet function, like (ctrl+q), it should be noted that if the user computer  ¦
'¦has different keyboard layouts like spanish or arabic differet shortcut keys need to be    ¦
'¦defined for the shortcut key in those layouts as well. In which case functions like        ¦
'¦password_View_Sheet_ or password_View_Sheet__ can be used.                                 ¦
'¦                                                                                           ¦
'¦   Such functionality can become handy in situations when a workbook has to be shared among¦
'¦multiple users and each user should have access to the sheets based on priorities or their ¦
'¦clearance and access level.                                                                ¦
'¦   It should be noted that the use of this method in highly sensitive environments is      ¦
'¦STRONGLY DISCOURAGED and it SHOULD NOT be used in those situations as the fundamental      ¦
'¦method used in Excel can be easily broken, and the sheet protection circumvented using     ¦
'¦brute force algorithms such as the famous:                                                 ¦
'¦                                                                                           ¦
'¦Sub PasswordBreaker()                                                                      ¦
'¦Dim i As Integer, j As Integer, k As Integer                                               ¦
'¦Dim l As Integer, m As Integer, n As Integer                                               ¦
'¦Dim i1 As Integer, i2 As Integer, i3 As Integer                                            ¦
'¦Dim i4 As Integer, i5 As Integer, i6 As Integer                                            ¦
'¦On Error Resume Next                                                                       ¦
'¦For i = 65 To 66: For j = 65 To 66: For k = 65 To 66                                       ¦
'¦For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66                                      ¦
'¦For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66                                    ¦
'¦For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126                                    ¦
'¦ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _                                         ¦
'¦Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _                                          ¦
'¦Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)                                                       ¦
'¦If ActiveSheet.ProtectContents = False Then                                                ¦
'¦MsgBox "One usable password is " & Chr(i) & Chr(j) & _                                     ¦
'¦Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _                                           ¦
'¦Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)                                             ¦
'¦Exit Sub                                                                                   ¦
'¦End If                                                                                     ¦
'¦Next: Next: Next: Next: Next: Next                                                         ¦
'¦Next: Next: Next: Next: Next: Next                                                         ¦
'¦End Sub                                                                                    ¦
'¦                                                                                           ¦
'¦   Although some hashing algorithms like SHA512 or simple KDF (Key Derivation Function) can¦
'¦be added to the algorithm to slow down the effect of bruteforce, but in the end it is the  ¦
'¦fundamental Excel worksheet protection algorithm that has to be fixed by Microsoft.        ¦
'¦                                                                                           ¦
'¦   Some hashing algorithms like MD5, SHA1, SHA256, SHA512 implemented in VBA are provied   ¦
'¦below in case you intend to add such functionality. The code has been copied from:         ¦
'¦https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/String_Hashing_in_VBA          ¦
'¦                                                                                           ¦
'¦Option Explicit                                                                            ¦
'¦                                                                                           ¦
'¦Sub Test()                                                                                 ¦
'¦    'run this to test md5, sha1, sha2/256, sha384, sha2/512 with salt, or sha2/512         ¦
'¦    Dim sIn As String, sOut As String, b64 As Boolean                                      ¦
'¦    Dim sH As String, sSecret As String                                                    ¦
'¦                                                                                           ¦
'¦    'insert the text to hash within the sIn quotes                                         ¦
'¦    'and for selected procedures a string for the secret key                               ¦
'¦    sIn = ""                                                                               ¦
'¦    sSecret = "" 'secret key for StrToSHA512Salt only                                      ¦
'¦                                                                                           ¦
'¦    'select as required                                                                    ¦
'¦    'b64 = False   'output hex                                                             ¦
'¦    b64 = True   'output base-64                                                           ¦
'¦                                                                                           ¦
'¦    'enable any one                                                                        ¦
'¦    'sH = MD5(sIn, b64)                                                                    ¦
'¦    'sH = SHA1(sIn, b64)                                                                   ¦
'¦    'sH = SHA256(sIn, b64)                                                                 ¦
'¦    'sH = SHA384(sIn, b64)                                                                 ¦
'¦    'sH = StrToSHA512Salt(sIn, sSecret, b64)                                               ¦
'¦    sH = SHA512(sIn, b64)                                                                  ¦
'¦                                                                                           ¦
'¦    'message box and immediate window outputs                                              ¦
'¦    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"                         ¦
'¦    MsgBox sH & vbNewLine & Len(sH) & " characters in length"                              ¦
'¦                                                                                           ¦
'¦    'de-comment this block to place the hash in first cell of sheet1                       ¦
'¦'    With ThisWorkbook.Worksheets("Sheet1").Cells(1, 1)                                    ¦
'¦'        .Font.Name = "Consolas"                                                           ¦
'¦'        .Select: Selection.NumberFormat = "@" 'make cell text                             ¦
'¦'        .Value = sH                                                                       ¦
'¦'    End With                                                                              ¦
'¦                                                                                           ¦
'¦End Sub                                                                                    ¦
'¦                                                                                           ¦
'¦Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = 0) As String           ¦
'¦    'Set a reference to mscorlib 4.0 64-bit                                                ¦
'¦                                                                                           ¦
'¦    'Test with empty string input:                                                         ¦
'¦    'Hex:   d41d8cd98f00...etc                                                             ¦
'¦    'Base-64: 1B2M2Y8Asg...etc                                                             ¦
'¦                                                                                           ¦
'¦    Dim oT As Object, oMD5 As Object                                                       ¦
'¦   Dim TextToHash() As Byte                                                                ¦
'¦    Dim bytes() As Byte                                                                    ¦
'¦                                                                                           ¦
'¦    Set oT = CreateObject("System.Text.UTF8Encoding")                                      ¦
'¦    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")       ¦
'¦                                                                                           ¦
'¦    TextToHash = oT.Getbytes_4(sIn)                                                        ¦
'¦    bytes = oMD5.ComputeHash_2((TextToHash))                                               ¦
'¦                                                                                           ¦
'¦    If bB64 = True Then                                                                    ¦
'¦       MD5 = ConvToBase64String(bytes)                                                     ¦
'¦    Else                                                                                   ¦
'¦       MD5 = ConvToHexString(bytes)                                                        ¦
'¦    End If                                                                                 ¦
'¦                                                                                           ¦
'¦    Set oT = Nothing                                                                       ¦
'¦    Set oMD5 = Nothing                                                                     ¦
'¦                                                                                           ¦
'¦End Function                                                                               ¦
'¦                                                                                           ¦
'¦Public Function SHA1(sIn As String, Optional bB64 As Boolean = 0) As String                ¦
'¦    'Set a reference to mscorlib 4.0 64-bit                                                ¦
'¦                                                                                           ¦
'¦    'Test with empty string input:                                                         ¦
'¦    '40 Hex:   da39a3ee5e6...etc                                                           ¦
'¦    '28 Base-64:   2jmj7l5rSw0yVb...etc                                                    ¦
'¦                                                                                           ¦
'¦    Dim oT As Object, oSHA1 As Object                                                      ¦
'¦    Dim TextToHash() As Byte                                                               ¦
'¦    Dim bytes() As Byte                                                                    ¦
'¦                                                                                           ¦
'¦    Set oT = CreateObject("System.Text.UTF8Encoding")                                      ¦
'¦    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")                   ¦
'¦                                                                                           ¦
'¦    TextToHash = oT.Getbytes_4(sIn)                                                        ¦
'¦    bytes = oSHA1.ComputeHash_2((TextToHash))                                              ¦
'¦                                                                                           ¦
'¦    If bB64 = True Then                                                                    ¦
'¦      SHA1 = ConvToBase64String(bytes)                                                     ¦
'¦    Else                                                                                   ¦
'¦      SHA1 = ConvToHexString(bytes)                                                        ¦
'¦    End If                                                                                 ¦
'¦                                                                                           ¦
'¦    Set oT = Nothing                                                                       ¦
'¦    Set oSHA1 = Nothing                                                                    ¦
'¦                                                                                           ¦
'¦End Function                                                                               ¦
'¦                                                                                           ¦
'¦Public Function SHA256(sIn As String, Optional bB64 As Boolean = 0) As String              ¦
'¦   'Set a reference to mscorlib 4.0 64-bit                                                 ¦
'¦                                                                                           ¦
'¦  'Test with empty string input:                                                           ¦
'¦  '64 Hex:   e3b0c44298f...etc                                                             ¦
'¦   '44 Base-64:   47DEQpj8HBSa+/...etc                                                     ¦
'¦                                                                                           ¦
'¦   Dim oT As Object, oSHA256 As Object                                                     ¦
'¦   Dim TextToHash() As Byte, bytes() As Byte                                               ¦
'¦                                                                                           ¦
'¦   Set oT = CreateObject("System.Text.UTF8Encoding")                                       ¦
'¦   Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")                ¦
'¦                                                                                           ¦
'¦   TextToHash = oT.Getbytes_4(sIn)                                                         ¦
'¦  bytes = oSHA256.ComputeHash_2((TextToHash))                                              ¦
'¦                                                                                           ¦
'¦   If bB64 = True Then                                                                     ¦
'¦      SHA256 = ConvToBase64String(bytes)                                                   ¦
'¦   Else                                                                                    ¦
'¦      SHA256 = ConvToHexString(bytes)                                                      ¦
'¦   End If                                                                                  ¦
'¦                                                                                           ¦
'¦   Set oT = Nothing                                                                        ¦
'¦   Set oSHA256 = Nothing                                                                   ¦
'¦                                                                                           ¦
'¦End Function                                                                               ¦
'¦                                                                                           ¦
'¦Public Function SHA384(sIn As String, Optional bB64 As Boolean = 0) As String              ¦
'¦    'Set a reference to mscorlib 4.0 64-bit                                                ¦
'¦                                                                                           ¦
'¦   'Test with empty string input:                                                          ¦
'¦   '96 Hex:   38b060a751ac...etc                                                           ¦
'¦   '64 Base-64:   OLBgp1GsljhM2T...etc                                                     ¦
'¦                                                                                           ¦
'¦    Dim oT As Object, oSHA384 As Object                                                    ¦
'¦    Dim TextToHash() As Byte, bytes() As Byte                                              ¦
'¦                                                                                           ¦
'¦    Set oT = CreateObject("System.Text.UTF8Encoding")                                      ¦
'¦    Set oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")               ¦
'¦                                                                                           ¦
'¦    TextToHash = oT.Getbytes_4(sIn)                                                        ¦
'¦   bytes = oSHA384.ComputeHash_2((TextToHash))                                             ¦
'¦                                                                                           ¦
'¦   If bB64 = True Then                                                                     ¦
'¦       SHA384 = ConvToBase64String(bytes)                                                  ¦
'¦    Else                                                                                   ¦
'¦       SHA384 = ConvToHexString(bytes)                                                     ¦
'¦    End If                                                                                 ¦
'¦                                                                                           ¦
'¦    Set oT = Nothing                                                                       ¦
'¦    Set oSHA384 = Nothing                                                                  ¦
'¦                                                                                           ¦
'¦End Function                                                                               ¦
'¦                                                                                           ¦
'¦Public Function SHA512(sIn As String, Optional bB64 As Boolean = 0) As String              ¦
'¦    'Set a reference to mscorlib 4.0 64-bit                                                ¦
'¦                                                                                           ¦
'¦    'Test with empty string input:                                                         ¦
'¦    '128 Hex:   cf83e1357eefb8bd...etc                                                     ¦
'¦    '88 Base-64:   z4PhNX7vuL3xVChQ...etc                                                  ¦
'¦                                                                                           ¦
'¦    Dim oT As Object, oSHA512 As Object                                                    ¦
'¦    Dim TextToHash() As Byte, bytes() As Byte                                              ¦
'¦                                                                                           ¦
'¦    Set oT = CreateObject("System.Text.UTF8Encoding")                                      ¦
'¦    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")               ¦
'¦                                                                                           ¦
'¦    TextToHash = oT.Getbytes_4(sIn)                                                        ¦
'¦    bytes = oSHA512.ComputeHash_2((TextToHash))                                            ¦
'¦                                                                                           ¦
'¦    If bB64 = True Then                                                                    ¦
'¦       SHA512 = ConvToBase64String(bytes)                                                  ¦
'¦    Else                                                                                   ¦
'¦       SHA512 = ConvToHexString(bytes)                                                     ¦
'¦   End If                                                                                  ¦
'¦                                                                                           ¦
'¦    Set oT = Nothing                                                                       ¦
'¦    Set oSHA512 = Nothing                                                                  ¦
'¦                                                                                           ¦
'¦End Function                                                                               ¦
'¦                                                                                           ¦
'¦Function StrToSHA512Salt(ByVal sIn As String, ByVal sSecretKey As String, _                ¦
'¦                          Optional ByVal b64 As Boolean = False) As String                 ¦
'¦   'Returns a sha512 STRING HASH in function name, modified by the parameter sSecretKey.   ¦
'¦   'This hash differs from that of SHA512 using the SHA512Managed class.                   ¦
'¦   'HMAC class inputs are hashed twice;first input and key are mixed before hashing,       ¦
'¦   'then the key is mixed with the result and hashed again.                                ¦
'¦                                                                                           ¦
'¦  Dim asc As Object, enc As Object                                                         ¦
'¦  Dim TextToHash() As Byte                                                                 ¦
'¦   Dim SecretKey() As Byte                                                                 ¦
'¦   Dim bytes() As Byte                                                                     ¦
'¦                                                                                           ¦
'¦  'Test results with both strings empty:                                                   ¦
'¦   '128 Hex:    b936cee86c9f...etc                                                         ¦
'¦   '88 Base-64:   uTbO6Gyfh6pd...etc                                                       ¦
'¦                                                                                           ¦
'¦   'create text and crypto objects                                                         ¦
'¦   Set asc = CreateObject("System.Text.UTF8Encoding")                                      ¦
'¦                                                                                           ¦
'¦   'Any of HMACSHAMD5,HMACSHA1,HMACSHA256,HMACSHA384,or HMACSHA512 can be used             ¦
'¦   'for corresponding hashes, albeit not matching those of Managed classes.                ¦
'¦   Set enc = CreateObject("System.Security.Cryptography.HMACSHA512")                       ¦
'¦                                                                                           ¦
'¦    'make a byte array of the text to hash                                                 ¦
'¦    bytes = asc.Getbytes_4(sIn)                                                            ¦
'¦    'make a byte array of the private key                                                  ¦
'¦    SecretKey = asc.Getbytes_4(sSecretKey)                                                 ¦
'¦    'add the private key property to the encryption object                                 ¦
'¦   enc.Key = SecretKey                                                                     ¦
'¦                                                                                           ¦
'¦   'make a byte array of the hash                                                          ¦
'¦    bytes = enc.ComputeHash_2((bytes))                                                     ¦
'¦                                                                                           ¦
'¦   'convert the byte array to string                                                       ¦
'¦    If b64 = True Then                                                                     ¦
'¦       StrToSHA512Salt = ConvToBase64String(bytes)                                         ¦
'¦    Else                                                                                   ¦
'¦       StrToSHA512Salt = ConvToHexString(bytes)                                            ¦
'¦    End If                                                                                 ¦
'¦                                                                                           ¦
'¦   'release object variables                                                               ¦
'¦    Set asc = Nothing                                                                      ¦
'¦   Set enc = Nothing                                                                       ¦
'¦                                                                                           ¦
'¦End Function                                                                               ¦
'¦                                                                                           ¦
'¦Private Function ConvToBase64String(vIn As Variant) As Variant                             ¦
'¦                                                                                           ¦
'¦    Dim oD As Object                                                                       ¦
'¦                                                                                           ¦
'¦    Set oD = CreateObject("MSXML2.DOMDocument")                                            ¦
'¦      With oD                                                                              ¦
'¦        .LoadXML "<root />"                                                                ¦
'¦       .DocumentElement.DataType = "bin.base64"                                            ¦
'¦        .DocumentElement.nodeTypedValue = vIn                                              ¦
'¦      End With                                                                             ¦
'¦    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")                        ¦
'¦                                                                                           ¦
'¦    Set oD = Nothing                                                                       ¦
'¦                                                                                           ¦
'¦End Function                                                                               ¦
'¦                                                                                           ¦
'¦Private Function ConvToHexString(vIn As Variant) As Variant                                ¦
'¦                                                                                           ¦
'¦    Dim oD As Object                                                                       ¦
'¦                                                                                           ¦
'¦    Set oD = CreateObject("MSXML2.DOMDocument")                                            ¦
'¦                                                                                           ¦
'¦      With oD                                                                              ¦
'¦        .LoadXML "<root />"                                                                ¦
'¦        .DocumentElement.DataType = "bin.Hex"                                              ¦
'¦        .DocumentElement.nodeTypedValue = vIn                                              ¦
'¦     End With                                                                              ¦
'¦   ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")                            ¦
'¦                                                                                           ¦
'¦    Set oD = Nothing                                                                       ¦
'¦                                                                                           ¦
'¦End Function                                                                               ¦
'¦                                                                                           ¦
'¦                                                                                           ¦
'¦                                                                                           ¦
'¦                                                                                           ¦
'¦                                                                                           ¦
'¦                                                                                           ¦
'¦NOTE:In case you need to get rid of the ¦s and spaces a simple regex like:                 ¦
'¦[replace ' *¦$'  with ''] and [replace '^'¦' with ''] will do it :)                        ¦
'+-------------------------------------------------------------------------------------------+



'////////////////////////////////////////////////////////////////////
'Password masked inputbox
'Allows you to hide characters entered in a VBA Inputbox.
'
'Code written by Daniel Klann
'March 2003
'Thanks to WideBoyDixon over at https://www.excelforum.com for helping
'to get this to work in office 2016
'////////////////////////////////////////////////////////////////////
'API functions to be used
#If VBA7 Then
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
#Else
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
#End If
'Constants to be used in our API functions
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0
Private hHook As LongPtr
Public Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
    Dim RetVal
    Dim strClassName As String, lngBuffer As Long
    If lngCode < HC_ACTION Then
        NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
        Exit Function
    End If
    strClassName = String$(256, " ")
    lngBuffer = 255
    If lngCode = HCBT_ACTIVATE Then    'A window has been activated
        RetVal = GetClassName(wParam, strClassName, lngBuffer)
        If Left$(strClassName, RetVal) = "#32770" Then  'Class name of the Inputbox
            'This changes the edit control so that it display the password character *.
            'You can change the Asc("*") as you please.
            SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
        End If
    End If
    'This line will ensure that any other hooks that may be in place are
    'called correctly.
    CallNextHookEx hHook, lngCode, wParam, lParam
End Function
Function InputBoxDK(Prompt, Title) As String
    Dim lngModHwnd As LongPtr, lngThreadID As Long
    lngThreadID = GetCurrentThreadId
    lngModHwnd = GetModuleHandle(vbNullString)
    hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
    InputBoxDK = InputBox(Prompt, Title)
    UnhookWindowsHookEx hHook
End Function
Public Sub password_View_Sheet()
Attribute password_View_Sheet.VB_ProcData.VB_Invoke_Func = "q\n14"
Dim Password_To_Protect As String
Dim Check_Pass As String
Dim no_of_rows As String
If ActiveSheet.ProtectContents = True Then
    Password_To_Protect = InputBoxDK("Enter password to view the content:", "Password to unprotect the sheet")
    If Password_To_Protect = "" Or Password_To_Protect = " " Then
        MsgBox ("No password was chosen!" + vbNewLine + "The sheet cannot be viewed!")
        Exit Sub
    End If
    On Error GoTo JmpErr
    ActiveSheet.Unprotect Password_To_Protect
    If ActiveSheet.ProtectContents = False Then
        no_of_rows = Replace(("1:" + Str(Sheet1.Rows.Count())), " ", "")
        Rows(no_of_rows).Select
        Selection.EntireRow.Hidden = False
        Selection.RowHeight = 15
    End If
Else
    Password_To_Protect = InputBoxDK("Enter password to protect view:", "Password to protect viewing the sheet")
    If Password_To_Protect = "" Or Password_To_Protect = " " Then
        MsgBox ("No password was chosen!" + vbNewLine + "The sheet will not be protected!")
        Exit Sub
    End If
    Check_Pass = InputBoxDK("Enter the password again:", "Password confirmation")
    If Check_Pass <> Password_To_Protect Then
        MsgBox "The passwords did not match!", vbOKOnly, "Password verification failed"
        Exit Sub
    End If
    
    
    no_of_rows = Replace(("1:" + Str(Sheet1.Rows.Count())), " ", "")
    Rows(no_of_rows).Select
    Selection.EntireRow.Hidden = True
    ActiveSheet.Protect Password_To_Protect, True, True, True, False, False, False, False, False, False, False, False, False, False, False, False
End If
Range("A1").Select
Exit Sub
JmpErr:
MsgBox ("The password was incorrect")
End Sub
Public Sub password_View_Sheet_()
Attribute password_View_Sheet_.VB_ProcData.VB_Invoke_Func = "Q\n14"
'This sub is only here so that a second shortcut key like for Capital Q (Ctrl+Shift/Capslock+Q) can also be introduced
password_View_Sheet
End Sub
Public Sub password_View_Sheet__()
Attribute password_View_Sheet__.VB_ProcData.VB_Invoke_Func = "?\n14"
'This sub is only here so that a third shortcut key like for the equal of Q in another keyboard layout can also be introduced
password_View_Sheet
End Sub



