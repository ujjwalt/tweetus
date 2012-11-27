Attribute VB_Name = "Account"
Public Function verify(usr As String, pass As String) As Boolean
Dim resp As New DOMDocument
resp.async = False
url = Twitter_module.get_it("http://twitter.com/account/verify_credentials.xml", usr, pass)
resp.loadXML (url)
If resp.Text = "/account/verify_credentials.xml Could not authenticate you." Then
verify = False
Exit Function
Else
verify = True
End If
End Function
Public Sub rate_limit()
Dim resprate As New DOMDocument, url As String
resprate.async = False
url = get_it("http://twitter.com/account/rate_limit_status.xml", usr, pass)
resprate.loadXML (url)
Main.lblrate.Caption = resprate.selectSingleNode("//hash/remaining-hits").Text + " hits remaining out of " _
+ resprate.selectSingleNode("//hash/hourly-limit").Text
End Sub
Public Sub update(field As String, values As String)
p = post_it("http://twitter.com/account/update_profile.xml?" + field + "=" + values, usr, pass)
If Mid(p, 40, 6) = "<user>" Then
MsgBox "Updated Succesfully"
Else
MsgBox "Some error - please try again after some time", vbCritical
End If
End Sub
