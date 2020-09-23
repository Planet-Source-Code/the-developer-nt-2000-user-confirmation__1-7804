Attribute VB_Name = "Password"
Option Explicit
Public Declare Function NetUserChangePassword Lib "netapi32.dll" (strDomainName As Any, strUserName As Any, strOldPassword As Any, strNewPassword As Any) As Long
Public Function CheckPassword(DomianName As String, UserName As String, oldPassword As String, Newpassword As String) As Boolean
Dim ReturnCode As Long
Dim lpwstrDomianName() As Byte
Dim lpwstrUserName() As Byte
Dim lpwstroldPassword() As Byte
Dim lpwstrNewpassword() As Byte
Dim iFileResult As Integer

'Set unicode strings to array
lpwstrDomianName = DomianName & vbNullChar
lpwstrUserName = UserName & vbNullChar
lpwstroldPassword = oldPassword & vbNullChar
lpwstrNewpassword = Newpassword & vbNullChar

' OK what happens is.....
'We try to change the NT password passing the given password and an empty string.
'NT server should not allow the empty string as the new password.
'So you can check the result to see if the password passed was correct.

'If you try to use this to crack people passwords, a word of warning... It will take for ever!
'But as a means as Identification works a treat!!

                                ReturnCode = NetUserChangePassword(lpwstrDomianName(0), lpwstrUserName(0), lpwstroldPassword(0), lpwstrNewpassword(0))
                                
                                Select Case ReturnCode
                                 Case 53
                                          MsgBox "The Net work path was not found", vbOKOnly + vbInformation, "Interrigation notice"
                                          CheckPassword = False
                                Case 86
                                          MsgBox "The specified network password is not correct", vbOKOnly + vbInformation, "Interrigation notice"
                                        
                                Case 2245
                                         
                                           ' frmpassword.txtPassword.Text = lpwstroldPassword
                                           ' MsgBox "Password Found !", vbInformation + vbOKOnly, "Cracked the Password!"
                                            CheckPassword = True
                                                                           
                                  Case 2221
                                  ' USER NOT FOUND
                                         MsgBox "The specified User did not exist, or could not be found.", vbOKOnly + vbInformation, "Interrigation notice"
                                         CheckPassword = False
                                 Case 1355
                                ' DOMAIN NOT FOUND
                                            MsgBox "The specified domian did not exist.", vbOKOnly + vbInformation, "Interrigation notice"
                                            CheckPassword = False
                                Case Else
                                    MsgBox "Unchecked know return code from Network interigation. #" & ReturnCode, vbInformation + vbOKOnly, "Network information"
                                    CheckPassword = False
                            End Select
                        DoEvents
End Function
