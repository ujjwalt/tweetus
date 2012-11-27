Attribute VB_Name = "test"
Public Sub testme(resp1 As String)
Dim fs, f, p As String
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.OpenTextFile("D:\twitter\test.xml", 2)
f.Write resp1
f.Close
End Sub
Function auth(strUsername, strPassword)
Dim fs, f, p
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile("D:\twitter\test.txt", 1)
    p = f.readall
    f.Close

    ' This is the function wicht does all the work.
    ' It uses XMLHTTP to post your message to Twitter..
    Dim objHTTP
    Set objHTTP = CreateObject("Microsoft.XMLHTTP")
    
        objHTTP.open "POST", "http://twitter.com/oauth/request_token", False, strUsername, strPassword
        objHTTP.send ""
        
        ' The function stores the Twitter response to the result of the function so you can use this later
        auth = objHTTP.responseText
        
    Set objHTTP = Nothing 'Release the object
    
End Function
