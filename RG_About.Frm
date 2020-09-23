VERSION 5.00
Begin VB.Form RG_About 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8385
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBlack 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   7410
      Left            =   195
      ScaleHeight     =   570
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   195
      Width           =   7995
      Begin VB.PictureBox picBlue 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   7020
         Left            =   195
         ScaleHeight     =   540
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   195
         Width           =   7605
         Begin VB.PictureBox picTop 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   3250
            Left            =   195
            ScaleHeight     =   250
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   555
            TabIndex        =   8
            Top             =   195
            Width           =   7215
            Begin VB.PictureBox picLogo 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.34
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1560
               Left            =   4875
               Picture         =   "RG_About.frx":0000
               ScaleHeight     =   116
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   155
               TabIndex        =   9
               Top             =   1495
               Width           =   2067
            End
            Begin VB.Label lblCompanyName 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "RandyGrams - (2009)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   416
               Left            =   1807
               TabIndex        =   16
               Top             =   117
               Width           =   3601
            End
            Begin VB.Label lblPlatform 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Windows Compatible"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   234
               Left            =   364
               TabIndex        =   15
               Top             =   1495
               Width           =   1911
            End
            Begin VB.Label lblVersion 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Version  1.2.3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   234
               Left            =   364
               TabIndex        =   14
               Top             =   1826
               Width           =   1222
            End
            Begin VB.Label lblCopyright 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Copyright:  RandyGrams (2009)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   234
               Left            =   364
               TabIndex        =   13
               Top             =   2157
               Width           =   2821
            End
            Begin VB.Label lblProductName 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "RandyGrams"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   32
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   728
               Left            =   1651
               TabIndex        =   12
               Top             =   585
               Width           =   3913
            End
            Begin VB.Label lblWrittenBy 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Written By:   Randy Giese"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   234
               Left            =   364
               TabIndex        =   11
               Top             =   2488
               Width           =   2327
            End
            Begin VB.Label lblEmail 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Email:  WebMaster@RandyGrams.Com"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   234
               Left            =   364
               TabIndex        =   10
               Top             =   2821
               Width           =   3523
            End
         End
         Begin VB.PictureBox picBottom 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   2470
            Left            =   195
            ScaleHeight     =   190
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   555
            TabIndex        =   4
            Top             =   3640
            Width           =   7215
            Begin VB.Label lblFree 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00000000&
               Height          =   1105
               Left            =   156
               TabIndex        =   7
               Top             =   546
               Width           =   6903
            End
            Begin VB.Label lblSend 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Please send any Comments, Questions or Suggestions to: WebMaster@RandyGrams.Com"
               ForeColor       =   &H00FF0000&
               Height          =   520
               Left            =   702
               TabIndex        =   6
               Top             =   1911
               Width           =   5798
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblE_Mail_Ware 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "This program is 'E-Mail-Ware'."
               ForeColor       =   &H00FF0000&
               Height          =   260
               Left            =   2028
               TabIndex        =   5
               Top             =   39
               Width           =   3146
            End
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H00E0E0E0&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   520
            Index           =   0
            Left            =   6214
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   6305
            Width           =   1196
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H00E0E0E0&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   520
            Index           =   1
            Left            =   195
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   6305
            Width           =   1196
         End
         Begin VB.Label lblRG_Com 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "www.RandyGrams.com"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   403
            Left            =   2002
            TabIndex        =   17
            Top             =   6370
            Width           =   3588
         End
      End
   End
End
Attribute VB_Name = "RG_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    The About screen displays information about this             ****
'   ****    RandyGrams program including the current Version number.     ****
'   ****                                                                 ****
'   ****    Please visit:  --->  www.RandyGrams.com  <---                ****
'   ****                                                                 ****
'   ****    for more free programs.                                      ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Sub cmdOK_Click(Index As Integer)
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Return (or Go) to the Main program.                          ****
'   ****                                                                 ****
'   *************************************************************************
'
    Unload Me

End Sub

Private Sub Form_Activate()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    A special  "Thank You"  to  "David Leatherman"  for          ****
'   ****    helping me get rid of an unneeded Timer.                     ****
'   ****                                                                 ****
'   ****    Thank You David Leatherman!                                  ****
'   ****                                                                 ****
'   *************************************************************************
'
    With cmdOK(1)
'       Locate the Cursor in the center of the "OK" button.
        SetCursorPos (Me.Left \ Screen.TwipsPerPixelX) + picBlack.Left + picBlue.Left + .Left + (.Width \ 2), (Me.Top \ Screen.TwipsPerPixelY) + picBlack.Top + picBlue.Top + .Top + (.Height \ 2)
        .SetFocus
    End With  'CmdOK(1)

End Sub

Private Sub Form_Load()

Dim lngTxtWidth  As Long
Dim strAboutYear   As String
'
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks                                                                Thanks
'   Thanks    Many, many Thanks to Roger Gilchrist, who not only went     Thanks
'   Thanks    way out of his way but also well above and beyond the       Thanks
'   Thanks    call of duty, to help someone he has never met!!!           Thanks
'   Thanks                                                                Thanks
'   Thanks    I owe you a great debt of gratitude Roger!                  Thanks
'   Thanks                                                                Thanks
'   Thanks    Please check out the best program I have ever seen on       Thanks
'   Thanks    PSC:                                                        Thanks
'   Thanks                                                                Thanks
'   Thanks    "CodeFixer"                                                 Thanks
'   Thanks    By:  Roger Gilchrist.                                       Thanks
'   Thanks                                                                Thanks
'   Thanks    The latest version as of this writing can be found at:      Thanks
'   Thanks                                                                Thanks
'   Thanks    http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=53297&lngWId=1
'   Thanks                                                                Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    strAboutYear = "RandyGrams (" & Year(Now) & ")"        'Put the Year inside the ()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Print the Current Year on the RG-Logo.                       ****
'   ****                                                                 ****
'   *************************************************************************
'
    With picLogo
        .CurrentX = (.ScaleWidth - .TextWidth(strAboutYear)) / 2
        .CurrentY = .ScaleHeight - 35
        picLogo.Print strAboutYear
    End With  'Logo

    Do
        picTop.Font = lblProductName.Font
        picTop.FontSize = lblProductName.FontSize
        lngTxtWidth = picTop.TextWidth(App.Title)

        If lngTxtWidth > picTop.Width - 6 Then
            lblProductName.FontSize = lblProductName.FontSize - 4
        End If
    Loop Until lngTxtWidth < picTop.Width - 6

    lblCopyright.Caption = "Copyright: " & strAboutYear    'Add the Year to the Copyright
    lblCompanyName.Caption = strAboutYear   'Add the Year to the Company Name
    lblProductName.Caption = App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblFree.Caption = "This program is 100% free.  Please feel free to pass it on to your" & vbNewLine & " friends and E-Mail aquaintances.  All I ask is that you send" & vbNewLine & " me a brief E-Mail to let me know what you think of:" & vbNewLine & App.Title & "!  Thank you!"

End Sub
