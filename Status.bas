Attribute VB_Name = "status"
Public ind As Integer
Public Function tweet(msg As String) As Boolean
Dim resp As New DOMDocument
If msg = "" Then

    ' If nothing is entered, show a message
    MsgBox "Blank Tweets aren't allowed"
    frmtweetbox.tweetbox.Text = "Type at least something for God's Sake"
    frmtweetbox.tweetbox.SelStart = 1
    frmtweetbox.tweetbox.SelLength = Len(frmtweetbox.tweetbox.Text)
    frmtweetbox.tweetbox.SetFocus

Else

    ' If there is some input from the user,
    ' then post it to twitter.
    resp.async = False
    p = Twitter_module.post_it("http://twitter.com/statuses/update.xml", usr, pass, msg)
    resp.loadXML (p)
    If resp.Text <> "/account/verify_credentials.xml Could not authenticate you." Then
    tweet = True
    Else
    tweet = False
    End If
    
End If
End Function
Public Sub delete(i As Integer)
q = MsgBox("Are you sure? There is no UNDO", vbYesNo, "Sure ???")
If q = vbYes Then
Dim currnd%, id$, p$
currnd = round \ 3
id = resp.selectSingleNode("//statuses/status[" + str((8 * (round - 1)) + ind) + "]/id").Text
p = post_it("http://twitter.com/statuses/destroy/" + id + ".xml", usr, pass, "DELETE")
If Mid(p, 40, 8) = "<status>" Then
MsgBox "Deleted Succesfully"
Else
MsgBox "Some error - please try again after some time", vbCritical
End If
Else
Exit Sub
End If
End Sub
Public Sub reply(i As Integer)
Dim f As New frmtweetbox
f.show
f.tweetbox.Text = "@" & resp.selectSingleNode("//statuses/status[" + str((8 * (round - 1)) + ind) + "]/user/screen_name").Text & " "
f.tweetbox.SetFocus
End Sub
Public Sub rt(i As Integer)
Dim f As New frmtweetbox
f.show
f.tweetbox.Text = "RT @" & resp.selectSingleNode("//statuses/status[" + str((8 * (round - 1)) + ind) + "]/user/screen_name").Text & " " & resp.selectSingleNode("//statuses/status[" + str((8 * (round - 1)) + ind) + "]/text").Text
f.tweetbox.SetFocus
End Sub
