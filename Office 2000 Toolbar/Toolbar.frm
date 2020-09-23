VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Office2000 Toolbar"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Help 
      Height          =   945
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   1667
      ButtonWidth     =   2514
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "Color"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help             "
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "What is this  "
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About             "
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Edit 
      Height          =   1890
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   3334
      ButtonWidth     =   4710
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "Color"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Copy                                        "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Paste                                       "
            ImageIndex      =   16
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cut                                          "
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Font Size                                 "
            ImageIndex      =   17
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Font Name                              "
            ImageIndex      =   17
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find                                         "
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Toolbar.frx":0000
         Left            =   1320
         List            =   "Toolbar.frx":000D
         TabIndex        =   7
         Text            =   "Times New Roman"
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Text            =   "10"
         Top             =   975
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar File 
      Height          =   1260
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   2223
      ButtonWidth     =   2381
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "Color"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New             "
            ImageIndex      =   13
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open            "
            ImageIndex      =   14
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save            "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit               "
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Menu 
      Left            =   3720
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1
      ImageHeight     =   1
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":0039
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   556
      ButtonWidth     =   873
      ButtonHeight    =   503
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "Menu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "File"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Style           =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList Color 
      Left            =   3120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":0091
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":05D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":0B19
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":105D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":15A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":1AE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":2029
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":256D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":2685
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":2799
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":28AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":2DF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":3335
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":3879
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":3DBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":4301
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Toolbar.frx":4845
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "Color"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.Toolbar Toolbar3 
      Height          =   345
      Left            =   4320
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "Color"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mfile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnew 
         Caption         =   "New"
      End
      Begin VB.Menu mopen 
         Caption         =   "Open"
      End
      Begin VB.Menu mm 
         Caption         =   "-"
      End
      Begin VB.Menu mexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x1, y1


Private Sub Edit_ButtonClick(ByVal Button As MSComctlLib.Button)
Edit.Visible = False
UpMenu

Select Case Button.Index

Case 1
MsgBox "Copy"
Case 2
MsgBox "Paste"
Case 3
MsgBox "Cut"
Case 6
MsgBox "Find"
End Select

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MenuV
UpMenu
End Sub

Private Sub Help_ButtonClick(ByVal Button As MSComctlLib.Button)
Help.Visible = False
UpMenu

Select Case Button.Index

Case 1
MsgBox "Help"
Case 2
MsgBox "What is this"
Case 3
MsgBox "   Hussain Al-Omran" & vbCrLf & "HUS_ME@Yahoo.com"

End Select

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

Case 3
MsgBox "New"
Case 4
MsgBox "Open"
Case 5
MsgBox "Save"
Case 6
MsgBox "Preview"
Case 7
MsgBox "Print"
Case 8
MsgBox "Find"
Case 9
MsgBox "What's this help"
End Select

End Sub

Private Sub Toolbar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

x1 = x
y1 = y
MenuV
UpMenu


End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button Then
If x > 200 Then Exit Sub

If Toolbar1.Left > -10 Then Toolbar1.Left = Toolbar1.Left - x1 + x
If Toolbar1.Top > -10 Then Toolbar1.Top = Toolbar1.Top - y1 + y

End If

End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Toolbar1.Left < 0 Then Toolbar1.Left = 0
If Toolbar1.Top < 0 Then Toolbar1.Top = 0

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

If Button.Caption = "File" Then
File.Move Toolbar2.Left + 300, Toolbar2.Top + Toolbar2.Height
File.Visible = True
Edit.Visible = False
Help.Visible = False
End If

If Button.Caption = "Edit" Then
Edit.Move Toolbar2.Left + 700, Toolbar2.Top + Toolbar2.Height
Edit.Visible = True
File.Visible = False
Help.Visible = False
End If

If Button.Caption = "Help" Then
Help.Move Toolbar2.Left + 1300, Toolbar2.Top + Toolbar2.Height
Help.Visible = True
File.Visible = False
Edit.Visible = False

End If

End Sub

Private Sub Toolbar2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
x1 = x
y1 = y
MenuV
UpMenu
End Sub

Private Sub Toolbar2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button Then
If x < 200 Then

If Toolbar2.Left > -10 Then Toolbar2.Left = Toolbar2.Left - x1 + x
If Toolbar2.Top > -10 Then Toolbar2.Top = Toolbar2.Top - y1 + y

End If
End If

End Sub

Private Sub Toolbar2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Toolbar2.Left < 0 Then Toolbar2.Left = 0
If Toolbar2.Top < 0 Then Toolbar2.Top = 0

End Sub

Private Sub File_ButtonClick(ByVal Button As MSComctlLib.Button)
File.Visible = False
UpMenu

Select Case Button.Index

Case 1
MsgBox "New"
Case 2
MsgBox "Open"
Case 3
MsgBox "Save"
Case 4
MsgBox "Exit"
End
End Select


End Sub


Public Sub UpMenu()

For i = 1 To Toolbar2.Buttons.Count
Toolbar2.Buttons(i).Value = tbrUnpressed
Next i

End Sub

Public Sub MenuV()
File.Visible = False
Edit.Visible = False
Help.Visible = False
End Sub


Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

Case 3
MsgBox "Copy"
Case 4
MsgBox "Undo"
Case 5
MsgBox "Spell Check"
Case 6
MsgBox "Paint"
Case 7
MsgBox "What's this Help"
Case 8
MsgBox "   Hussain Al-Omran" & vbCrLf & "HUS_ME@Yahoo.com"

End Select

End Sub

Private Sub Toolbar3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
x1 = x
y1 = y
MenuV
UpMenu
End Sub

Private Sub Toolbar3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button Then
If x > 200 Then Exit Sub

If Toolbar3.Left > -10 Then Toolbar3.Left = Toolbar3.Left - x1 + x
If Toolbar3.Top > -10 Then Toolbar3.Top = Toolbar3.Top - y1 + y

End If

End Sub

Private Sub Toolbar3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Toolbar3.Left < 0 Then Toolbar3.Left = 0
If Toolbar3.Top < 0 Then Toolbar3.Top = 0

End Sub
