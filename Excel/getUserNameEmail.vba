Sub GetMicrosoftAccountName()
    Dim microsoftAccountName As String
    
    ' Retrieve the Microsoft account name
    microsoftAccountName = Application.UserName
    ' Display the Microsoft account name in a message box
    MsgBox "Microsoft Account Name: " & microsoftAccountName
End Sub

Sub getLoggedInUser()

    Dim olApp As Object
    Dim olNS As Object

    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")

    With olNS.session.currentuser.AddressEntry.GetExchangeUser
        MsgBox (.FirstName & " " & .LastName) ' This is the name
        MsgBox .primarysmtpaddress ' This is the primary email
    End With

End Sub
