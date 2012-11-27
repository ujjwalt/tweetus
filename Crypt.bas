Attribute VB_Name = "Crypt"
Public Function Hash(pass As String, usr As String) As String
Dim Array_pass(1 To 50) As String, revpass As String, asc(1 To 50), ascpass As String
revpass = pass

For i = 1 To Len(pass)
Array_pass(i) = Mid(pass, Len(pass) - (i - 1), 1)
revpass = revpass + Array_pass(i)
Next i

For j = 1 To Len(revpass)
ascpass = AscW(Mid(revpass, j, 1)) & ascpass
Next j

Hash = md5(Salt(usr) & ascpass)
End Function
Public Function Salt(user As String)
Salt = md5(user)
End Function
Public Function md5(str As String)

    Dim objHTTP, resp, resp_pos
    Set objHTTP = CreateObject("Microsoft.XMLHTTP")
    
        objHTTP.open "GET", "http://www.iwebtool.com/tool/tools/md5/md5.php" + "?string=" + str, False
        objHTTP.send "nothing"
        
        ' The function stores the Twitter response to the result of the function so you can use this later
        resp = objHTTP.responseText
        resp_pos = Mid(resp, 81 + Len(str), 32)
        
    Set objHTTP = Nothing 'Release the object
    md5 = resp_pos
End Function
