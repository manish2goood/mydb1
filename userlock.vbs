' ResetPasswordAndLogout.vbs
' This script resets a user's password on a Windows machine and logs out the current user

Option Explicit

Dim username, password, objShell, objNetwork, objUser

' Prompt for username and new password
username = InputBox("Enter the username to reset the password:", "INTG")
password = InputBox("Enter the new password:", "123456")

' Create WScript.Shell object to run commands
Set objShell = CreateObject("WScript.Shell")

' Create WScript.Network object to access network properties
Set objNetwork = CreateObject("WScript.Network")

' Get the domain name
Dim domain
domain = objNetwork.UserDomain

' Bind to the user object in Active Directory
On Error Resume Next
Set objUser = GetObject("WinNT://" & domain & "/" & username & ",user")

' Reset the password
If Not objUser Is Nothing Then
    objUser.SetPassword password
    If Err.Number = 0 Then
        MsgBox "Password for user " & username & " has been reset successfully.", vbInformation, "Success"
    Else
        MsgBox "Failed to reset the password for user " & username & ": " & Err.Description, vbCritical, "Error"
    End If
Else
    MsgBox "User " & username & " not found.", vbCritical, "Error"
End If

' Clean up
Set objUser = Nothing
Set objNetwork = Nothing

' Log out the current user
objShell.Run "shutdown -l", 0, False

' Clean up shell object
Set objShell = Nothing
