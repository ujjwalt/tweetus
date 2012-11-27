Attribute VB_Name = "Twitter_module"
Public usr As String, pass As String, round As Integer
Public Function Twitter(status) As Boolean
Twitter = False
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.OpenTextFile("D:\twitter\status_error.txt", 1)
p = f.readall
f.Close
For i = 1 To Len(status)
For j = 1 To i
If Mid(status, j, i) = p Then
MsgBox "Invalid Username or Password - check again plz !", vbOKOnly, "TWITTER STATUS UPDATE"
frmLogin.show
frmLogin.txtPassword.Text = ""
frmLogin.txtUserName.SetFocus
Exit Function
Else
Twitter = True
End If
Next j
Next i
End Function
Function post_it(strurl As String, Optional strUsername As String, Optional strPassword As String, Optional strMessage As String) As String
If strurl <> "http://twitter.com/account/verify_credentials.xml" Or strurl <> "http://twitter.com/account/rate_limit_status.xml" Then
Call rate_limit
End If
    ' This is the function wicht does all the work.
    ' It uses XMLHTTP to post your message to Twitter..
    Dim objHTTP
    Set objHTTP = CreateObject("Microsoft.XMLHTTP")
    
        objHTTP.open "POST", strurl, False, strUsername, strPassword
        objHTTP.send "status=" & strMessage
        
        ' The function stores the Twitter response to the result of the function so you can use this later
        post_it = objHTTP.responseText
        
    Set objHTTP = Nothing 'Release the object
    
End Function
Function get_it(strurl As String, Optional strUsername As String, Optional strPassword As String) As String
If strurl <> "http://twitter.com/account/verify_credentials.xml" And strurl <> "http://twitter.com/account/rate_limit_status.xml" Then
Call rate_limit
End If
    ' This is the function wicht does all the work.
    ' It uses XMLHTTP to post your message to Twitter..
    Dim objHTTP
    Set objHTTP = CreateObject("Microsoft.XMLHTTP")
    
        objHTTP.open "GET", strurl, False, strUsername, strPassword
        objHTTP.send ""
        
        ' The function stores the Twitter response to the result of the function so you can use this later
        get_it = objHTTP.responseText
        Debug.Print strurl
        
    Set objHTTP = Nothing 'Release the object
    
End Function
