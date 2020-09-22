VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleMode       =   0  'User
   ScaleWidth      =   9630.38
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgView 
      Height          =   495
      Index           =   1
      Left            =   3000
      Top             =   6480
      Width           =   3615
   End
   Begin VB.Image imgView 
      Height          =   495
      Index           =   0
      Left            =   3000
      Top             =   6480
      Width           =   3615
   End
   Begin VB.Image imgMain 
      Height          =   855
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Image imgTouch 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   0
      MouseIcon       =   "frmMain.frx":0000
      MousePointer    =   99  'Custom
      ToolTipText     =   "Click Me!"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image imgTouch 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   0
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      ToolTipText     =   "Click Me!"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape shpTouch 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   1200
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape shpTouch 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   4695
      Index           =   0
      Left            =   2880
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Image imgMain 
      Height          =   855
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Image imgMain 
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Image imgMain 
      Height          =   855
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Image imgBack 
      Height          =   495
      Left            =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgExamples 
      Height          =   975
      Left            =   0
      Picture         =   "frmMain.frx":0614
      Top             =   3600
      Width           =   2250
   End
   Begin VB.Image imgExit 
      Height          =   630
      Left            =   8400
      Picture         =   "frmMain.frx":2AA4
      Top             =   6600
      Width           =   1200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'We'll make some arrays to handle some of the graphics
Dim Quit(4) As String
Dim Back(2) As String
Dim Example(2) As String

Private Sub Form_Load()

'here we populate the arrays with the graphic files
Quit(0) = App.Path & "\graphics\exitup.jpg"
Quit(1) = App.Path & "\graphics\exitdn.jpg"
Quit(2) = App.Path & "\graphics\earthexitup.jpg"
Quit(3) = App.Path & "\graphics\earthexitdn.jpg"

Back(0) = App.Path & "\graphics\earthbackup.jpg"
Back(1) = App.Path & "\graphics\earthbackdn.jpg"

Example(0) = App.Path & "\graphics\exampleup.jpg"
Example(1) = App.Path & "\graphics\exampledn.jpg"

'Here we've made control arrays to handle the graphics
imgView(0).Picture = LoadPicture(App.Path & "\graphics\viewup.jpg")
imgView(0).Visible = False
imgView(1).Picture = LoadPicture(App.Path & "\graphics\hideup.jpg")
imgView(1).Visible = False

With imgMain(0)
    .Picture = LoadPicture(App.Path & "\graphics\back.jpg")
    .Visible = True
    .ZOrder 1
    With imgMain(1)
        .Picture = LoadPicture(App.Path & "\graphics\earthbg.jpg")
        .Visible = False
        .ZOrder 1
        With imgMain(2)
            .Picture = LoadPicture(App.Path & "\graphics\moon.jpg")
            .Visible = False
            .ZOrder 1
            With imgMain(3)
                .Picture = LoadPicture(App.Path & "\graphics\earth.jpg")
                .Visible = False
                .ZOrder 1
            End With
        End With
    End With
End With

'Let's hide the hotspot locators
shpTouch(0).Visible = False
shpTouch(1).Visible = False

End Sub

Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'We want to change the graphic
imgBack.Picture = LoadPicture(Back(1))

End Sub

Private Sub imgBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Here we set up the screen.., it's basic but it gets the job done
With imgExamples
    .Picture = LoadPicture(Example(0))
    .Visible = False

    With imgBack
        .Picture = LoadPicture(Back(0))
        .Visible = True
                    
        With imgTouch(0)
            .Enabled = True
            .Top = shpTouch(0).Top
            .Left = shpTouch(0).Left
            .Width = shpTouch(0).Width
            .Height = shpTouch(0).Height
                
            With imgTouch(1)
                .Enabled = True
                .Top = shpTouch(1).Top
                .Left = shpTouch(1).Left
                .Width = shpTouch(1).Width
                .Height = shpTouch(1).Height
                
                With imgView(0)
                    .Enabled = True
                    .Visible = True
                    
                End With
            End With
        End With
    End With
End With

imgMain(1).Visible = True

imgMain(0).Visible = False

imgMain(2).Visible = False

imgMain(3).Visible = False

imgExit.Picture = LoadPicture(Quit(2))


End Sub

Private Sub imgExamples_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgExamples.Picture = LoadPicture(Example(1))

End Sub

Private Sub imgExamples_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'This sets up the screen for the Example
With imgExamples
    .Picture = LoadPicture(Example(0))
    .Visible = False

    With imgBack
        .Picture = LoadPicture(Back(0))
        .Visible = True
                    
        With imgTouch(0)
            .Enabled = True
            .Top = shpTouch(0).Top
            .Left = shpTouch(0).Left
            .Width = shpTouch(0).Width
            .Height = shpTouch(0).Height
                
            With imgTouch(1)
                .Enabled = True
                .Top = shpTouch(1).Top
                .Left = shpTouch(1).Left
                .Width = shpTouch(1).Width
                .Height = shpTouch(1).Height
                
            End With
        End With
    End With
End With

imgMain(1).Visible = True

imgMain(0).Visible = False

imgExit.Picture = LoadPicture(Quit(2))

imgView(0).Visible = True

frmExplain.Show

End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Here we have to find out what screen is visible and what exit button to present
If imgMain(1).Visible Then
    imgExit.Picture = LoadPicture(Quit(3))
ElseIf imgMain(2).Visible Then
    imgExit.Picture = LoadPicture(Quit(3))
ElseIf imgMain(3).Visible Then
    imgExit.Picture = LoadPicture(Quit(3))
Else
    imgExit.Picture = LoadPicture(Quit(1))
End If


End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MsgBox "Thank you for viewing the HotSpot program" & vbCrLf & _
     "     For comments or suggestions contact" & vbCrLf & _
     "         hogana@usapathway.com"
    End
    
End Sub


Private Sub imgTouch_Click(Index As Integer)

'Which hot spot is the user choosing and what is the screen apperance?
Select Case Index

    Case 0
        imgMain(3).Visible = True
        imgMain(1).Visible = False
        imgTouch(0).Enabled = False
        imgTouch(1).Enabled = False
        shpTouch(0).Visible = False
        shpTouch(1).Visible = False
        imgView(0).Enabled = False
        imgView(0).Visible = False
        imgView(1).Enabled = False
        imgView(1).Visible = False
        
    Case 1
        imgMain(2).Visible = True
        imgMain(1).Visible = False
        imgTouch(0).Enabled = False
        imgTouch(1).Enabled = False
        shpTouch(0).Visible = False
        shpTouch(1).Visible = False
        imgView(0).Enabled = False
        imgView(0).Visible = False
        imgView(1).Enabled = False
        imgView(1).Visible = False
End Select

End Sub

Private Sub imgView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'We want to find out if the touch points are visible or not so we can present the
'correct view or hide
If imgView(0).Visible = True Then
    imgView(0).Picture = LoadPicture(App.Path & "\graphics\viewdn.jpg")
Else
    imgView(1).Picture = LoadPicture(App.Path & "\graphics\hidedn.jpg")
End If

End Sub

Private Sub imgView_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Now we have to reset the buttons back
If imgView(0).Visible = True Then
    imgView(1).Visible = True
    shpTouch(0).Visible = True
    shpTouch(1).Visible = True
    imgView(0).Picture = LoadPicture(App.Path & "\graphics\viewup.jpg")
    imgView(0).Visible = False
Else
    imgView(1).Visible = True
    shpTouch(0).Visible = False
    shpTouch(1).Visible = False
    imgView(0).Picture = LoadPicture(App.Path & "\graphics\viewup.jpg")
    imgView(0).Visible = True
    imgView(1).Visible = False
    imgView(1).Picture = LoadPicture(App.Path & "\graphics\hideup.jpg")
End If

End Sub
