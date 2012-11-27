VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.UserControl Tweets 
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   ScaleHeight     =   1335
   ScaleWidth      =   5535
   Begin RichTextLib.RichTextBox inv 
      Height          =   135
      Left            =   4800
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   238
      _Version        =   393217
      TextRTF         =   $"Tweet.ctx":0000
   End
   Begin RichTextLib.RichTextBox msg 
      Height          =   1095
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1931
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"Tweet.ctx":008B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblshow 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lbldel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      ToolTipText     =   "Delete this tweet"
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblrt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "RT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "ReTweet"
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblrep 
      BackColor       =   &H00E0E0E0&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Reply to this tweet"
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblfav 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Favourite this"
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image pp 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   0
      Stretch         =   -1  'True
      Top             =   360
      Width           =   720
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Tweets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dis As Boolean
Public Property Get tweettxt() As String
tweettxt = msg.Text
End Property
Public Property Let tweettxt(str As String)
msg.Text = str
ln = Len(msg.Text)
msg.SelStart = 0
msg.SelLength = ln
msg.SelColor = vbBlack
msg.SelLength = 0

    If ln > 1 Then
        For i = 1 To ln
            inv.TextRTF = msg.TextRTF
            inv.SelStart = i
                If Mid(inv.Text, i, 1) = "#" Or Mid(inv.Text, i, 1) = "@" Or Mid(inv.Text, i, 7) = "http://" Then
                    For j = i To ln
                        If Mid(inv.Text, j, 1) = Space(1) Or j = ln Then
                        inv.SelStart = i - 1
                        inv.SelLength = j - i + 1
                        inv.SelColor = vbBlue
                        inv.SelStart = j - 1
                        inv.SelColor = vbBlack
                        msg.TextRTF = inv.TextRTF
                        Exit For
                        End If
                    Next j
                End If
        Next i
    End If
End Property
Public Property Let ppa(url As String)
On Error GoTo errors
If url = "already loaded" Then
pp.Picture = LoadPicture("C:\ppimg." + frmt)
Exit Property
End If
Dim bData() As Byte
bData() = Main.Inet1.OpenURL(url, icByteArray)

Select Case Mid(url, Len(url) - 2, 3)
Case "png"
Open "C:\pp.png" For Binary Access Write As #1
Put #1, , bData()
Close #1
pp.Picture = LoadPicture("C:\pp.png")
Case "jpg", "peg"
Open "C:\pp.jpg" For Binary Access Write As #1
Put #1, , bData()
Close #1
pp.Picture = LoadPicture("C:\pp.jpg")
Case Else
Open "C:\pp.bmp" For Binary Access Write As #1
Put #1, , bData()
Close #1
pp.Picture = LoadPicture("C:\pp.bmp")
End Select
Exit Property
errors:
pp.ToolTipText = "Due to some error at Twitter - profile pic couldn't be loaded"
pp.Picture = LoadPicture("C:\ppp.jpg")
Exit Property
End Property
Public Property Let tooltip(txt As String)
If pp.ToolTipText = "Due to some error at Twitter - profile pic couldn't be loaded" Then
pp.ToolTipText = pp.ToolTipText + Space(1) + "-" + Space(1) + txt
Else
pp.ToolTipText = txt
End If
End Property
Private Sub lbldel_Click()
status.delete (ind)
invis
End Sub
Private Sub lblfav_Click()
invis
End Sub

Private Sub lblrep_Click()
Call status.reply(ind)
End Sub

Private Sub lblrt_Click()
Call status.rt(ind)
End Sub

Private Sub lblshow_Click()
Call user.show(ind)
End Sub

Private Sub pp_Click()
If dis = False Then
lblfav.Visible = Not lblfav.Visible
lblrep.Visible = Not lblrep.Visible
lbldel.Visible = Not lbldel.Visible
lblrt.Visible = Not lblrt.Visible
lblshow.Visible = Not lblshow.Visible
End If
End Sub

Private Sub invis()
lblfav.Visible = False
lblrep.Visible = False
lbldel.Visible = False
lblrt.Visible = False
lblshow.Visible = False
End Sub
Public Property Let disable(i As Boolean)
dis = i
End Property


