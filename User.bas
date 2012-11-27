Attribute VB_Name = "User"
Public yes As Boolean
Public Sub show(Optional i As Integer = -1)
yes = True
On Error GoTo err
If i = -1 Then
gouser = InputBox("Enter Screen Name")
Else
gouser = resp.selectSingleNode("//statuses/status[" + str(i) + "]/user/screen_name").Text
End If

If gouser <> "" Then
Dim f As New frmUser
Load f
If gouser = usr Then
With f
.txtdesc.Enabled = True
.txtloc.Enabled = True
.txtname.Enabled = True
.txturl.Enabled = True
End With
Else
yes = False
End If
p = get_it("http://twitter.com/users/show/" + gouser + ".xml", usr, pass)
Dim locresp As New DOMDocument
locresp.async = False
locresp.loadXML (p)
Dim bData() As Byte
bData() = Main.Inet1.OpenURL(locresp.selectSingleNode("//user/profile_image_url").Text, icByteArray)

Open "C:\ppuser.bmp" For Binary Access Write As #1
Put #1, , bData()
Close #1
f.pp.Picture = LoadPicture("C:\ppuser.bmp")

f.lbldesc = locresp.selectSingleNode("//user/description").Text
If locresp.selectSingleNode("//user/following").Text = "true" Then
f.lblfol = "Yes you're following " + locresp.selectSingleNode("//user/name").Text
Else
f.lblfol = "Not Following"
f.cmdfollow.Enabled = True
f.cmdfollow.Visible = True
End If
f.lblname = locresp.selectSingleNode("//user/name").Text
f.Caption = f.Caption & f.lblname
f.lblloc = locresp.selectSingleNode("//user/location").Text
f.lblstat = locresp.selectSingleNode("//user/status/text").Text
f.lbldesc = locresp.selectSingleNode("//user/description").Text
f.lblurl = locresp.selectSingleNode("//user/url").Text
f.show
Exit Sub
Else
MsgBox "Enter a Name"
Exit Sub
End If
err:
f.pp.Picture = LoadPicture("C:\ppp.jpg")
End Sub
