Attribute VB_Name = "Get_Timeline"
Public resp As New DOMDocument, respuser As New DOMDocument, frmt As String, rounduser As Integer, name1 As String, name2 As String, roundme As Integer
Public Sub timeline(Optional prev As Boolean)
Static minid(202), resplen As Integer, id
For i = 4 To 202 Step 3
    Select Case round
        Case i, 1
            If resp.Text <> "" And prev = False Then
            minid(round - 1) = id
                id = Val(resp.selectSingleNode("//statuses/status[0]/id").Text)
                    For m = 0 To resp.childNodes.Item(1).childNodes.Length - 1
                        maxid = Val(resp.selectSingleNode("//statuses/status[" + str(m) + "]/id").Text)
                        
                            If maxid < id Then
                                id = maxid
                                End If
                                Next m
                                resp.async = False
                                url = Twitter_module.get_it("http://twitter.com/statuses/friends_timeline.xml?max_id=" & id, usr, pass)
                                resp.loadXML (url)
                                GoTo 15
                                ElseIf prev = True Then
                                For k = 0 To 7
                                Main.msg(k).Visible = True
                                Main.msg(k + 8).Visible = False
                                Next k
                                ElseIf round = 1 And prev = False Then
                                resp.async = False
                        url = Twitter_module.get_it("http://twitter.com/statuses/friends_timeline.xml", usr, pass)
                        resp.loadXML (url)
15
                        For k = 0 To 7
                                Main.msg(k).Visible = True
                                Main.msg(k + 8).Visible = False
                                Main.msg(k \ 2 + 16).Visible = False
                                Next k
                                Main.Hide
                       frmSplash.show
                       frmSplash.pb.Max = resp.childNodes.Item(1).childNodes.Length - 1
                       resplen = frmSplash.pb.Max
                       For j = 0 To resplen
                    frmSplash.pb.Value = j
                    Main.msg(j).ppa = resp.selectSingleNode("//statuses/status[" + str(j) + "]/user/profile_image_url").Text
                    Main.msg(j).tooltip = resp.selectSingleNode("//statuses/status[" + str(j) + "]/user/name").Text
                    Main.msg(j).tweettxt = resp.selectSingleNode("//statuses/status[" + str(j) + "]/text").Text
                    Next j
                    Unload frmSplash
                    Main.show
                    End If
        Exit Sub
        
        Case i + 1, 2
        If resplen < 15 Then
                    resplen = resp.childNodes.Item(1).childNodes.Length - 1

                    Else
                    resplen = 15
                    End If
16
                    If resplen < 8 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To resplen
                    Main.msg(j).Visible = False
                    Next j
        For j = 8 To resplen
        Main.msg(j).Visible = True
        Next j
        Exit Sub
        
        Case i + 2, 3
        If prev = True Then
                        resp.async = False
                       If round <> 3 Then
                        url = Twitter_module.get_it("http://twitter.com/statuses/friends_timeline.xml?max_id=" & minid(round), usr, pass)
                        Else
                        url = Twitter_module.get_it("http://twitter.com/statuses/friends_timeline.xml", usr, pass)
                       End If
                       resp.loadXML (url)
                        Main.Hide
                       frmSplash.show
                       frmSplash.pb.Max = resp.childNodes.Item(1).childNodes.Length - 1
                       resplen = frmSplash.pb.Max
                        For j = 0 To resplen
                    frmSplash.pb.Value = j
                    Main.msg(j).ppa = resp.selectSingleNode("//statuses/status[" + str(j) + "]/user/profile_image_url").Text
                    Main.msg(j).tooltip = resp.selectSingleNode("//statuses/status[" + str(j) + "]/user/name").Text
                    Main.msg(j).tweettxt = resp.selectSingleNode("//statuses/status[" + str(j) + "]/text").Text
                    Next j
                    Unload frmSplash
                        End If
                        
        If resplen < 19 Then
                    resplen = resp.childNodes.Item(1).childNodes.Length - 1
  
                    Else
                    resplen = 19
                    End If
                    
                    If resplen < 8 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To resplen
                    Main.msg(j).Visible = False
                    Next j
                    
        For j = 16 To resplen
        Main.msg(j).Visible = True
        Next j
        Main.show
        Exit Sub
    End Select
Next i

End Sub
Public Sub usertime(Optional name1 As String = "", Optional prev As Boolean)
On Error GoTo fin
If name1 = "" Then
If name2 = "" Then
name2 = InputBox("Specify a name")
If name2 = "" Then
Exit Sub
End If
End If
End If
Static minid(202), resplen As Integer, id
For i = 4 To 202 Step 3
    Select Case rounduser
        Case i, 1
            If respuser.Text <> "" And prev = False Then
            minid(rounduser - 1) = id
                id = Val(respuser.selectSingleNode("//statuses/status[0]/id").Text)
                    For m = 0 To respuser.childNodes.Item(1).childNodes.Length - 1
                        maxid = Val(respuser.selectSingleNode("//statuses/status[" + str(m) + "]/id").Text)
                        
                            If maxid < id Then
                                id = maxid
                                End If
                                Next m
                                respuser.async = False
                                url = Twitter_module.get_it("http://twitter.com/statuses/user_timeline/" + name2 + ".xml?max_id=" & id, usr, pass)
                                respuser.loadXML (url)
                                GoTo 15
                                ElseIf prev = True Then
                                For k = 0 To 7
                                MainUser.msg(k).Visible = True
                                MainUser.msg(k + 8).Visible = False
                                Next k
                                ElseIf rounduser = 1 And prev = False Then
                                respuser.async = False
                        url = Twitter_module.get_it("http://twitter.com/statuses/user_timeline/" + name2 + ".xml", usr, pass)
                        respuser.loadXML (url)
15
                                MainUser.Hide
                                image (respuser.selectSingleNode("//statuses/status[0]/user/profile_image_url").Text)
                       frmSplash.show
                       frmSplash.pb.Max = respuser.childNodes.Item(1).childNodes.Length - 1
                       resplen = frmSplash.pb.Max
                       For j = 0 To resplen
                    frmSplash.pb.Value = j
                    MainUser.msg(j).ppa = "already loaded"
                    MainUser.msg(j).tooltip = respuser.selectSingleNode("//statuses/status[" + str(j) + "]/user/name").Text
                    MainUser.msg(j).tweettxt = respuser.selectSingleNode("//statuses/status[" + str(j) + "]/text").Text
                    Next j
                                            For k = 0 To 7
                                MainUser.msg(k).Visible = True
                                MainUser.msg(k + 8).Visible = False
                                MainUser.msg(k \ 2 + 16).Visible = False
                                Next k
                    Unload frmSplash
                    MainUser.show
                    End If
        Exit Sub
        
        Case i + 1, 2
        If resplen < 15 Then
                    resplen = respuser.childNodes.Item(1).childNodes.Length - 1

                    Else
                    resplen = 15
                    End If
16
                    If resplen < 8 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To resplen
                    MainUser.msg(j).Visible = False
                    Next j
        For j = 8 To resplen
        MainUser.msg(j).Visible = True
        Next j
        Exit Sub
        
        Case i + 2, 3
        If prev = True Then
                        respuser.async = False
                       If rounduser <> 3 Then
                        url = Twitter_module.get_it("http://twitter.com/statuses/user_timeline/" + name2 + ".xml?max_id=" & minid(rounduser), usr, pass)
                        Else
                        url = Twitter_module.get_it("http://twitter.com/statuses/user_timeline/" + name2 + ".xml", usr, pass)
                       End If
                       respuser.loadXML (url)
                        MainUser.Hide
                       frmSplash.show
                       frmSplash.pb.Max = respuser.childNodes.Item(1).childNodes.Length - 1
                       resplen = frmSplash.pb.Max
                        For j = 0 To resplen
                    frmSplash.pb.Value = j
                    MainUser.msg(j).ppa = "already loaded"
                    MainUser.msg(j).tooltip = respuser.selectSingleNode("//statuses/status[" + str(j) + "]/user/name").Text
                    MainUser.msg(j).tweettxt = respuser.selectSingleNode("//statuses/status[" + str(j) + "]/text").Text
                    Next j
                    Unload frmSplash
                        End If
                        
        If resplen < 19 Then
                    resplen = respuser.childNodes.Item(1).childNodes.Length - 1
  
                    Else
                    resplen = 19
                    End If
                   If resplen < 16 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To resplen
                    MainUser.msg(j).Visible = False
                    Next j
                    
        For j = 16 To resplen
        MainUser.msg(j).Visible = True
        Next j
        MainUser.show
        Exit Sub
    End Select
Next i
fin:
Unload MainUser
Unload frmSplash
name1 = ""
name2 = ""
id = 0
resplen = 0
rounduser = 1
respuser.loadXML ("")
End Sub
Public Sub Mentions(Optional prev As Boolean)
'On Error GoTo fin
Static minid(202), resplen As Integer, id, respme As New DOMDocument
For i = 4 To 202 Step 3
    Select Case roundme
        Case i, 1
            If respme.Text <> "" And prev = False Then
            minid(roundme - 1) = id
                id = Val(respme.selectSingleNode("//statuses/status[0]/id").Text)
                    For m = 0 To respme.childNodes.Item(1).childNodes.Length - 1
                        maxid = Val(respme.selectSingleNode("//statuses/status[" + str(m) + "]/id").Text)
                        
                            If maxid < id Then
                                id = maxid
                                End If
                                Next m
                                respme.async = False
                                url = Twitter_module.get_it("http://twitter.com/statuses/mentions.xml", usr, pass)
                                respme.loadXML (url)
                                GoTo 15
                                ElseIf prev = True Then
                                For k = 0 To 7
                                Mention.msg(k).Visible = True
                                Mention.msg(k + 8).Visible = False
                                Next k
                                ElseIf roundme = 1 And prev = False Then
                                respme.async = False
                        url = Twitter_module.get_it("http://twitter.com/statuses/mentions.xml", usr, pass)
                        respme.loadXML (url)
15
                                Mention.Hide
                       frmSplash.show
                       frmSplash.pb.Max = respme.childNodes.Item(1).childNodes.Length - 1
                       resplen = frmSplash.pb.Max
                       For j = 0 To resplen
                    frmSplash.pb.Value = j
                    Mention.msg(j).ppa = respme.selectSingleNode("//statuses/status[" + str(j) + "]/user/profile_image_url").Text
                    Mention.msg(j).tooltip = respme.selectSingleNode("//statuses/status[" + str(j) + "]/user/name").Text
                    Mention.msg(j).tweettxt = respme.selectSingleNode("//statuses/status[" + str(j) + "]/text").Text
                    Next j
                                            For k = 0 To 7
                                Mention.msg(k).Visible = True
                                Mention.msg(k + 8).Visible = False
                                Mention.msg(k \ 2 + 16).Visible = False
                                Next k
                    Unload frmSplash
                    Mention.show
                    End If
        Exit Sub
        
        Case i + 1, 2
        If resplen < 15 Then
                    resplen = respme.childNodes.Item(1).childNodes.Length - 1

                    Else
                    resplen = 15
                    End If
16
                    If resplen < 8 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To resplen
                    Mention.msg(j).Visible = False
                    Next j
        For j = 8 To resplen
        Mention.msg(j).Visible = True
        Next j
        Exit Sub
        
        Case i + 2, 3
        If prev = True Then
                        respme.async = False
                       If roundme <> 3 Then
                        url = Twitter_module.get_it("http://twitter.com/statuses/mentions.xml" & minid(roundme), usr, pass)
                        Else
                        url = Twitter_module.get_it("http://twitter.com/statuses/mentions.xml", usr, pass)
                       End If
                       respme.loadXML (url)
                        Mention.Hide
                       frmSplash.show
                       frmSplash.pb.Max = respme.childNodes.Item(1).childNodes.Length - 1
                       resplen = frmSplash.pb.Max
                        For j = 0 To resplen
                    frmSplash.pb.Value = j
                    Mention.msg(j).ppa = "already loaded"
                    Mention.msg(j).tooltip = respme.selectSingleNode("//statuses/status[" + str(j) + "]/user/name").Text
                    Mention.msg(j).tweettxt = respme.selectSingleNode("//statuses/status[" + str(j) + "]/text").Text
                    Next j
                    Unload frmSplash
                        End If
                        
        If resplen < 19 Then
                    resplen = respme.childNodes.Item(1).childNodes.Length - 1
  
                    Else
                    resplen = 19
                    End If
                   If resplen < 16 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To resplen
                    Mention.msg(j).Visible = False
                    Next j
                    
        For j = 16 To resplen
        Mention.msg(j).Visible = True
        Next j
        Mention.show
        Exit Sub
    End Select
Next i
fin:
Unload Mention
Unload frmSplash
id = 0
resplen = 0
roundme = 1
respme.loadXML ("")
End Sub
Private Sub image(url As String)

Dim bData() As Byte
bData() = Main.Inet1.OpenURL(url, icByteArray)

Select Case Mid(url, Len(url) - 2, 3)
Case "png"
Open "C:\ppimg.png" For Binary Access Write As #1
Put #1, , bData()
frmt = "png"
Close #1
Case "jpg", "peg"
Open "C:\ppimg.jpg" For Binary Access Write As #1
Put #1, , bData()
frmt = "jpg"
Close #1
Case Else
Open "C:\ppimg.bmp" For Binary Access Write As #1
Put #1, , bData()
frmt = "bmp"
Close #1
End Select

End Sub
