Attribute VB_Name = "direct"
Public roundm As Integer, roundmsent As Integer
Public Sub directmsg(Optional prev As Boolean, Optional reset As Boolean)
On Error GoTo fin
If reset Then
GoTo fin
End If
Static minid(202), id, respdm As New DOMDocument, respdmlen As Integer
For i = 4 To 202 Step 3
    Select Case roundm
        Case i, 1
            If respdm.Text <> "" And prev = False Then
            For s = 0 To 19
            DM.msg(s).Visible = False
            DMSent.msg(s).tweettxt = ""
            Next s
            minid(roundm - 1) = id
                id = Val(respdm.selectSingleNode("//direct-messages/direct_message[0]/id").Text)
                    For m = 0 To respdm.childNodes.Item(1).childNodes.Length - 1
                        maxid = Val(respdm.selectSingleNode("//direct-messages/direct_message[" + str(m) + "]/id").Text)
                        
                            If maxid < id Then
                                id = maxid
                                End If
                                Next m
                                respdm.async = False
                                url = Twitter_module.get_it("http://twitter.com/direct_messages.xml?max_id=" & id, usr, pass)
                                respdm.loadXML (url)
                                GoTo 15
                                ElseIf prev = True Then
                                For k = 0 To 7
                                DM.msg(k).Visible = True
                                DM.msg(k + 8).Visible = False
                                Next k
                                ElseIf roundm = 1 And prev = False Then
                                respdm.async = False
                        url = Twitter_module.get_it("http://twitter.com/direct_messages.xml", usr, pass)
                        respdm.loadXML (url)
15
                        For k = 0 To 7
                                DM.msg(k).Visible = True
                                DM.msg(k + 8).Visible = False
                                DM.msg(k \ 2 + 16).Visible = False
                                Next k
                                DM.Hide
                       frmSplash.show
                       frmSplash.pb.Max = respdm.childNodes.Item(1).childNodes.Length - 1
                       respdmlen = frmSplash.pb.Max
                       For j = 0 To respdmlen
                    frmSplash.pb.Value = j
                    DM.msg(j).ppa = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/sender/profile_image_url").Text
                    DM.msg(j).tooltip = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/sender/name").Text & " to " _
                    & respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/recipient/name").Text
                    DM.msg(j).tweettxt = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/text").Text
                    Next j
                    Unload frmSplash
                    DM.show
                    End If
        Exit Sub
        
        Case i + 1, 2
        If respdmlen < 15 Then
                    respdmlen = respdm.childNodes.Item(1).childNodes.Length - 1

                    Else
                    respdmlen = 15
                    End If
16
                    If respdmlen < 8 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To respdmlen
                    DM.msg(j).Visible = False
                    Next j
        For j = 8 To respdmlen
        DM.msg(j).Visible = True
        Next j
        Exit Sub
        
        Case i + 2, 3
        If prev = True Then
                        respdm.async = False
                       If roundm <> 3 Then
                        url = Twitter_module.get_it("http://twitter.com/direct_messages.xml?max_id=" & minid(roundm), usr, pass)
                        Else
                        url = Twitter_module.get_it("http://twitter.com/direct_messages.xml", usr, pass)
                       End If
                       respdm.loadXML (url)
                        DM.Hide
                       frmSplash.show
                       frmSplash.pb.Max = respdm.childNodes.Item(1).childNodes.Length - 1
                       respdmlen = frmSplash.pb.Max
                        For j = 0 To respdmlen
                    frmSplash.pb.Value = j
                    DM.msg(j).ppa = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/sender/profile_image_url").Text
                    DM.msg(j).tooltip = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/sender/name").Text & " to " _
                    & respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/recipient/name").Text
                    DM.msg(j).tweettxt = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/text").Text
                    Next j
                    Unload frmSplash
                        End If
                        
        If respdmlen < 19 Then
                    respdmlen = respdm.childNodes.Item(1).childNodes.Length - 1
  
                    Else
                    respdmlen = 19
                    End If
                    
                    If respdmlen < 8 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To respdmlen
                    DM.msg(j).Visible = False
                    Next j
                    
        For j = 16 To respdmlen
        DM.msg(j).Visible = True
        Next j
        DM.show
        Exit Sub
    End Select
Next i
Exit Sub
fin:
respdmlen = 0
id = 0
respdm.loadXML ("")
roundm = roundm - 1
Unload frmSplash
directmsg
Unload DM
End Sub
Public Sub sent(Optional prev As Boolean, Optional reset As Boolean)
On Error GoTo fin
If reset Then
GoTo fin
End If
Static minid(202), respdmlen1 As Integer, id, respdm As New DOMDocument
For i = 4 To 202 Step 3
    Select Case roundmsent
        Case i, 1
            If respdm.Text <> "" And prev = False Then
            For s = 0 To 19
            DMSent.msg(s).Visible = False
            DMSent.msg(s).tweettxt = ""
            Next s
            minid(roundmsent - 1) = id
                id = Val(respdm.selectSingleNode("//direct-messages/direct_message[0]/id").Text)
                    For m = 0 To respdm.childNodes.Item(1).childNodes.Length - 1
                        maxid = Val(respdm.selectSingleNode("//direct-messages/direct_message[" + str(m) + "]/id").Text)
                        
                            If maxid < id Then
                                id = maxid
                                End If
                                Next m
                                respdm.async = False
                                url = Twitter_module.get_it("http://twitter.com/direct_messages/sent.xml?max_id=" & id, usr, pass)
                                respdm.loadXML (url)
                                GoTo 15
                                ElseIf prev = True Then
                                For k = 0 To 7
                                DMSent.msg(k).Visible = True
                                DMSent.msg(k + 8).Visible = False
                                Next k
                                ElseIf roundmsent = 1 And prev = False Then
                                Load DMSent
                                respdm.async = False
                        url = Twitter_module.get_it("http://twitter.com/direct_messages/sent.xml", usr, pass)
                        respdm.loadXML (url)
15
                        For k = 0 To 7
                                DMSent.msg(k).Visible = True
                                DMSent.msg(k + 8).Visible = False
                                DMSent.msg(k \ 2 + 16).Visible = False
                                Next k
                                DMSent.Hide
                       frmSplash.show
                       frmSplash.pb.Max = respdm.childNodes.Item(1).childNodes.Length - 1
                       respdmlen1 = respdm.childNodes.Item(1).childNodes.Length - 1
                       For j = 0 To respdmlen1
                    frmSplash.pb.Value = j
                    DMSent.msg(j).ppa = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/sender/profile_image_url").Text
                    DMSent.msg(j).tooltip = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/sender/name").Text & " to " _
                    & respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/recipient/name").Text
                    DMSent.msg(j).tweettxt = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/text").Text
                    Next j
                    Unload frmSplash
                    DMSent.show
                    End If
        Exit Sub
        
        Case i + 1, 2
        If respdmlen1 < 15 Then
                    respdmlen1 = respdm.childNodes.Item(1).childNodes.Length - 1

                    Else
                    respdmlen1 = 15
                    End If
16
                    If respdmlen1 < 8 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To respdmlen1
                    DMSent.msg(j).Visible = False
                    Next j
        For j = 8 To respdmlen1
        DMSent.msg(j).Visible = True
        Next j
        Exit Sub
        
        Case i + 2, 3
        If prev = True Then
                        respdm.async = False
                       If roundmsent <> 3 Then
                        url = Twitter_module.get_it("http://twitter.com/direct_messages/sent.xml?max_id=" & minid(roundmsent), usr, pass)
                        Else
                        url = Twitter_module.get_it("http://twitter.com/direct_messages/sent.xml", usr, pass)
                       End If
                       respdm.loadXML (url)
                        DMSent.Hide
                       frmSplash.show
                       frmSplash.pb.Max = respdm.childNodes.Item(1).childNodes.Length - 1
                       respdmlen1 = frmSplash.pb.Max
                        For j = 0 To respdmlen1
                    frmSplash.pb.Value = j
                    DMSent.msg(j).ppa = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/sender/profile_image_url").Text
                    DMSent.msg(j).tooltip = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/sender/name").Text & " to " _
                    & respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/recipient/name").Text
                    DMSent.msg(j).tweettxt = respdm.selectSingleNode("//direct-messages/direct_message[" + str(j) + "]/text").Text
                    Next j
                    Unload frmSplash
                        End If
                        
        If respdmlen1 < 19 Then
                    respdmlen1 = respdm.childNodes.Item(1).childNodes.Length - 1
  
                    Else
                    respdmlen1 = 19
                    End If
                    
                    If respdmlen1 < 8 Then
                    MsgBox "Nothing more to serve you"
                    Exit Sub
                    End If
                    For j = 0 To respdmlen1
                    DMSent.msg(j).Visible = False
                    Next j
                    
        For j = 16 To respdmlen1
        DMSent.msg(j).Visible = True
        Next j
        DMSent.show
        Exit Sub
    End Select
Next i
Exit Sub
fin:
respdmlen1 = 0
id = 0
respdm.loadXML ("")
Unload frmSplash
Unload DMSent
End Sub
Public Function newmsg()
Dim f As New frmtweetbox
f.show
f.tweetbox.Text = "D " & resp.selectSingleNode("//statuses/status[" + str((8 * (round - 1)) + ind) + "]/user/screen_name").Text & " "
f.tweetbox.SetFocus
End Function
Public Function destroy()

End Function
