VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormatterMain 
   BackColor       =   &H80000016&
   Caption         =   "Formatter     (www.RandyGrams.com)"
   ClientHeight    =   5525
   ClientLeft      =   104
   ClientTop       =   455
   ClientWidth     =   9126
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Formatter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   702
   Begin VB.PictureBox picTopBotFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2847
      Left            =   130
      ScaleHeight     =   219
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   628
      TabIndex        =   9
      Top             =   130
      Width           =   8164
      Begin VB.PictureBox picNumSpaces 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   858
         Left            =   3744
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   330
         TabIndex        =   17
         Top             =   975
         Width           =   4290
         Begin VB.TextBox txtNumSpaces 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   28
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   624
            Left            =   3601
            MaxLength       =   2
            TabIndex        =   18
            Text            =   "4"
            ToolTipText     =   "Enter the number of spaces between the Border and the Text!"
            Top             =   117
            Width           =   650
         End
         Begin VB.Label lblNumSpaces 
            BackColor       =   &H80000016&
            Caption         =   "Enter the number of spaces before the Asterisks (*).      This includes the Tick mark (')."
            ForeColor       =   &H00000000&
            Height          =   780
            Left            =   117
            TabIndex        =   19
            Top             =   39
            Width           =   3172
         End
      End
      Begin VB.PictureBox picTopBotBars 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   858
         Left            =   130
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   268
         TabIndex        =   14
         Top             =   975
         Width           =   3484
         Begin VB.TextBox txtTopBotBars 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   28
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   624
            Left            =   2795
            MaxLength       =   2
            TabIndex        =   15
            Text            =   "2"
            ToolTipText     =   "Enter the number of Top and Bottom Border lines!"
            Top             =   117
            Width           =   650
         End
         Begin VB.Label lblTopBotBars 
            BackColor       =   &H80000016&
            Caption         =   "Enter the number of Top and Bottom Border Bars!"
            ForeColor       =   &H00000000&
            Height          =   520
            Left            =   130
            TabIndex        =   16
            Top             =   169
            Width           =   2691
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox txtSideFormat 
         BackColor       =   &H00E2EFF1&
         BeginProperty Font 
            Name            =   "OCRB"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   130
         TabIndex        =   13
         Text            =   "'   ****                                                                 ****"
         Top             =   2353
         Width           =   7904
      End
      Begin VB.TextBox txtTopBotFormat 
         BackColor       =   &H00E2EFF1&
         BeginProperty Font 
            Name            =   "OCRB"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   130
         TabIndex        =   10
         Text            =   "'   *************************************************************************"
         Top             =   65
         Width           =   7904
      End
      Begin VB.Label lblSideFormat 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Caption         =   "Enter the format for the sides of the text box below."
         ForeColor       =   &H00000000&
         Height          =   286
         Left            =   1222
         TabIndex        =   12
         Top             =   1950
         Width           =   5720
      End
      Begin VB.Label lblTopBotFormat 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Caption         =   "Enter the Top and Bottom line format above."
         ForeColor       =   &H00000000&
         Height          =   286
         Left            =   1222
         TabIndex        =   11
         Top             =   572
         Width           =   5720
      End
   End
   Begin VB.PictureBox picEnterComments 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1430
      Left            =   130
      ScaleHeight     =   110
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   685
      TabIndex        =   6
      Top             =   3965
      Width           =   8905
      Begin VB.TextBox txtEnterComments 
         BackColor       =   &H00E2EFF1&
         BeginProperty Font 
            Name            =   "OCRB"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   715
         Left            =   468
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   20
         Text            =   "Formatter.frx":0CCA
         Top             =   468
         Width           =   5512
      End
      Begin VB.CheckBox chkLeftAlign 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Add - ""Left Align"" statement."
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   5252
         TabIndex        =   7
         Top             =   0
         Width           =   3419
      End
      Begin VB.Label lblEnterComments 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Enter your comments below, then press 'Format'!"
         ForeColor       =   &H00000000&
         Height          =   260
         Left            =   78
         TabIndex        =   8
         Top             =   0
         Width           =   4966
      End
   End
   Begin VB.PictureBox picCommandBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   754
      Left            =   130
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   525
      TabIndex        =   0
      Top             =   3094
      Width           =   6825
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H80000016&
         Caption         =   "Clear!"
         Height          =   754
         Left            =   0
         MaskColor       =   &H80000007&
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "  Erase  A L L  text!  "
         Top             =   0
         Width           =   1001
      End
      Begin VB.CommandButton cmdFormat 
         BackColor       =   &H80000016&
         Caption         =   "Format!"
         Height          =   754
         Left            =   1118
         MaskColor       =   &H80000007&
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "This will Format the Text using the current parameters."
         Top             =   0
         Width           =   1196
      End
      Begin VB.CommandButton cmdCopy 
         BackColor       =   &H80000016&
         Caption         =   "Copy!"
         Height          =   754
         Left            =   2431
         MaskColor       =   &H80000007&
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Copy the Text to the ClipBoard!"
         Top             =   0
         Width           =   1001
      End
      Begin VB.CommandButton cmdSaveText 
         BackColor       =   &H80000016&
         Caption         =   "Save Text!"
         Height          =   754
         Left            =   3549
         MaskColor       =   &H80000007&
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Temporarily Save the Text!"
         Top             =   0
         Width           =   1430
      End
      Begin VB.CommandButton cmdRestoreText 
         BackColor       =   &H80000016&
         Caption         =   "Restore Text!"
         Height          =   754
         Left            =   5096
         MaskColor       =   &H80000007&
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Restore the Text!"
         Top             =   0
         Width           =   1729
      End
   End
   Begin MSComDlg.CommonDialog CDB 
      Left            =   8073
      Top             =   3094
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuUnFormat 
         Caption         =   "Un-Format"
      End
   End
   Begin VB.Menu mnuFormatter 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuSample 
      Caption         =   "Show Sample"
   End
End
Attribute VB_Name = "FormatterMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Formatter                                                    ****
'   ****                                                                 ****
'   ****    Written by:    Randy Giese    (12/10/02)                     ****
'   ****                                                                 ****
'   ****    This program will build comment fields like this one so      ****
'   ****    that you may put them into your own programs. Run this       ****
'   ****    program and type or paste anything you like in the Text      ****
'   ****    Box.                                                         ****
'   ****                                                                 ****
'   ****    If you want a phrase or sentence to be on it's own line,     ****
'   ****    just press 'Enter' at the end of the line.                   ****
'   ****    For example:                                                 ****
'   ****    I pressed 'Enter' after the word "line." and again after     ****
'   ****    the word "example:" in the 2 lines above.  That will make    ****
'   ****    the line "For example:" be on a line by itself.              ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    December 05, 2009                                            ****
'   ****                                                                 ****
'   ****    I changed the "ScaleMode" property on the Form and all       ****
'   ****    Picture Boxes from Twips to Pixels.                          ****
'   ****                                                                 ****
'   ****    I modified the Form Layout slightly and changed the          ****
'   ****    Maximum Length of the Format Bars from 99 to 500.            ****
'   ****                                                                 ****
'   ****    I also changed the Rich Text Box to a plain Text Box         ****
'   ****    since no RTB features were being utilized.                   ****
'   ****                                                                 ****
'   ****    I did a little minor "Code Cleanup" and removed the          ****
'   ****    temporary Text file.                                         ****
'   ****                                                                 ****
'   ****    I now check for the Screen Width in Pixels.  If the          ****
'   ****    Screen Width is <1100 Pixels, I change the Font and the      ****
'   ****    Size of the 2 labels that are centered within the Gray       ****
'   ****    box above.  That gives sufficient room to correctly          ****
'   ****    display the 2 Text boxes that are also in the Gray box       ****
'   ****    above.                                                       ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    December 15, 2009                                            ****
'   ****                                                                 ****
'   ****    I used the "Split" function to replace a large segment of    ****
'   ****    "If Then" statements that used "InStr" and "InStrRev" to     ****
'   ****    find and remove all of the Asterisks "*" and Spaces " "      ****
'   ****    from the line of data.  Split is much quicker and more       ****
'   ****    efficient both "Timewise" and "Spacewise".                   ****
'   ****                                                                 ****
'   *************************************************************************
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    If you leave a Blank line in your text, it will leave a      ****
'   ****    Blank line in the formatted output as well, like this:       ****
'   ****                                                                 ****
'   ****    And, if you want to have a Breaker bar in your formatted     ****
'   ****    output, leave 2 Blank lines.                                 ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    Each consecutive Blank line will add another Breaker bar     ****
'   ****    in your formatted output.                                    ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Multiple Breaker bars are displayed above.                   ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Below is the text used to generate the previous Formatted    ****
'   ****    output.  You may try it yourself by selecting and copying    ****
'   ****    the text field below.  Then run the program and paste the    ****
'   ****    text into the text box.  Remove the tick marks (') but do    ****
'   ****    not delete any blank lines.  Then click Format.              ****
'   ****                                                                 ****
'   ****    The Formatted text should look like the text above here.     ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'
'   If you leave a Blank line in your text, it will leave a Blank line in the formatted output as well, like this:
'
'   And, if you want to have a Breaker bar in your formatted output, leave 2 Blank lines.
'
'
'   Each consecutive Blank line will add another Breaker bar in your formatted output.
'
'
'
'
'   Multiple Breaker bars are displayed above.
'
'
'
'   RandyGrams Comments - Left Align the above comments.

'Enter the number of spaces between the Border and the Text!
Option Explicit

Private lngBorderEnd         As Long
Private lngCounter           As Long
Private lngFieldEnd          As Long
Private lngFieldSize         As Long
Private lngLeftBorder        As Long
Private lngLoop2             As Long
Private lngScreenWide        As Long
Private lngSequentialCRLF    As Long
Private lngSpaces            As Long

Private strArray()           As String
Private strBreakerBar        As String
Private strOpenFile          As String
Private strSave              As String
Private strSideFormatBar     As String
Private strTopFormatBar      As String

'Define Constants
Private Const SM_CXSCREEN    As Long = 0    'Width of screen
Private Const SW_MAXIMIZE    As Long = 3

Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()

Public Sub BreakerBar()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This will determine if a Breaker Bar is needed or if it's    ****
'   ****    just a blank line.                                           ****
'   ****                                                                 ****
'   ****    If lngSequentialCRLF = 1 then we just print a blank line.    ****
'   ****    If it's > 1, then we print a Breaker Bar.                    ****
'   ****                                                                 ****
'   *************************************************************************
'
    If LenB(strArray(lngCounter)) = 0 Then
        lngSequentialCRLF = lngSequentialCRLF + 1

        If lngSequentialCRLF > 1 Then
'           This will cause the program to print a Breaker Bar.
            strArray(lngCounter) = "Place a BreakerBar here!"
        End If
    Else
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    If this line has text but lngSequentialCRLF > 1 then         ****
'   ****    insert a Blank line into the array.                          ****
'   ****                                                                 ****
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

        If lngSequentialCRLF > 1 Then
            lngCounter = lngCounter + 1     'This will be the number of the current array string
            strArray(lngCounter) = strArray(lngCounter - 1)               'Move the current array string data to the next array number
            strArray(lngCounter - 1) = vbNullString        'Insert a Blank line
        End If

        lngSequentialCRLF = 0
    End If

End Sub

Private Sub cmdClear_Click()

    txtEnterComments.Text = vbNullString
    txtEnterComments.SetFocus

End Sub

Private Sub cmdCopy_Click()

    With txtEnterComments
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

    Clipboard.Clear
    Clipboard.SetText txtEnterComments.SelText

End Sub

Private Sub cmdFormat_Click()

Dim lngCRLF     As Long
Dim lngLoop1    As Long
Dim strComments As String
Dim strCopy     As String
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The user clicked the "Format" button.                        ****
'   ****                                                                 ****
'   *************************************************************************
'
    ReDim strArray(0 To 999) As String      'Clear the Array
'   Store the Format Bars
    strTopFormatBar = txtTopBotFormat.Text
    strSideFormatBar = txtSideFormat.Text

    GetFormatBarInfo

'   "lngFieldSize" is the length of a string that will fit exactly between the "*" Borders.
    lngFieldSize = InStrRev(strSideFormatBar, " ", , vbTextCompare) - lngBorderEnd - (lngSpaces * 2)
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Clean any blank lines or spaces off the back of the text.    ****
'   ****                                                                 ****
'   *************************************************************************
'
    strComments = Trim$(txtEnterComments.Text)             'Trim any Spaces off the front and back of the string.
    lngCRLF = 1              'Make sure we go through the While/Wend at least once
    lngCounter = Len(strComments) - 1       'Where to start searching strComments for a CR/LF

    Do While lngCRLF > 0
        lngCRLF = InStr(lngCounter, strComments, vbNewLine, vbTextCompare)               'Check the end of the string for a CR/LF
        If lngCRLF > 0 Then
            strComments = Trim$(Left$(strComments, lngCounter - 1))       'Trim any Spaces and CR/LF off the back of the string.
            lngCounter = Len(strComments) - 1              'Where to start searching strComments for a CR/LF
        End If
    Loop

    strComments = Trim$(strComments) & vbNewLine           'Make sure that there are no Spaces at the end of the string.
    lngCounter = 0           'Clear the array Counter
    lngSequentialCRLF = 0

strCheck_It:
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Break the Text string down into individual sentences.        ****
'   ****                                                                 ****
'   *************************************************************************
'
    lngCRLF = InStr(1, strComments, vbLf, vbTextCompare)   'Check the string for a CR/LF

    If lngCRLF > 0 Then
        lngCounter = lngCounter + 1         'This will be the number of the current array string
        strArray(lngCounter) = Left$(strComments, lngCRLF - 2)            'Fill the array string up to the 1st line feed
        strCopy = Trim$(Mid$(strComments, lngCRLF + 1, 9999))             'Chop the CR/LF off the end of the string
    End If

    BreakerBar

'   Reset the string so we can continue checking it.
    If Len(strCopy) > 0 Then
        strComments = strCopy
        GoTo strCheck_It
    End If

'   If the last line is a Breaker, delete it.
    If strArray(lngCounter) = "Place a BreakerBar here!" Then
        strArray(lngCounter) = vbNullString
        lngCounter = lngCounter - 1
    End If
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Display the Formatted Comments.                              ****
'   ****                                                                 ****
'   *************************************************************************
'
    txtEnterComments.Text = vbNewLine & "'"

    If txtTopBotBars.Text > 0 Then
        For lngLoop1 = 1 To txtTopBotBars.Text             'Print the Top Border
            txtEnterComments.Text = txtEnterComments.Text & vbNewLine & strTopFormatBar
        Next lngLoop1
    End If

    txtEnterComments.Text = txtEnterComments.Text & vbNewLine & strSideFormatBar
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Break the text into strings that are 53 characters long.     ****
'   ****    Note: This number will change depending on the format        ****
'   ****    entered by the user.                                         ****
'   ****                                                                 ****
'   *************************************************************************
'
    For lngLoop1 = 1 To lngCounter
        strComments = Trim$(strArray(lngLoop1))            'strComments is the current comment

        If LenB(strComments) = 0 Then
            strBreakerBar = vbNullString
            GoTo PadTheString
        End If

        Do While Len(strComments) > 0
            strBreakerBar = Trim$(Left$(strComments, lngFieldSize + 1))   'strBreakerBar is the current comment segment

            If Len(strBreakerBar) < lngFieldSize + 1 Then
                lngLoop2 = Len(strBreakerBar) + 1
                GoTo PadTheString
            End If

            For lngLoop2 = Len(strBreakerBar) To 1 Step -1 'Start at the back of the line
                If Mid$(strBreakerBar, lngLoop2, 1) = " " Then            'Look for the first " " (Space character)
                    Exit For 'When you find it, exit the loop
                End If
            Next lngLoop2

PadTheString:
            If lngLoop2 = 0 Then
                lngLoop2 = lngFieldSize + 1
            End If

            PadLine

            strComments = Trim$(Mid$(strComments, lngLoop2, 9999))        'strComments is the remainder of the comment
        Loop
    Next lngLoop1
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Print the end lines out to the temporary file.               ****
'   ****                                                                 ****
'   *************************************************************************
'
    txtEnterComments.Text = txtEnterComments.Text & vbNewLine & strSideFormatBar

    If txtTopBotBars.Text > 0 Then
        For lngLoop1 = 1 To txtTopBotBars.Text             'Print the Top Border
            txtEnterComments.Text = txtEnterComments.Text & vbNewLine & strTopFormatBar
        Next lngLoop1
    End If

    txtEnterComments.Text = txtEnterComments.Text & vbNewLine & "'"

    If chkLeftAlign.Value = 1 Then
        txtEnterComments.Text = txtEnterComments.Text & vbNewLine & "'   RandyGrams Comments - Left Align the above comments."
    End If

End Sub

Private Sub cmdRestoreText_Click()

    txtEnterComments.Text = strSave

End Sub

Private Sub cmdSaveText_Click()

    strSave = txtEnterComments.Text

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

'   Store the Default Format Bars
    strTopFormatBar = "'   *************************************************************************"
    strSideFormatBar = "'   ****                                                                 ****"

    GetFormatBarInfo

'   Get the Screen Width
    lngScreenWide = GetSystemMetrics(SM_CXSCREEN)
'   Maximize the Form
    Me.Move 0, 0, Screen.Width, Screen.Height
'   Display a Sample
    txtEnterComments.Text = "This is an example of what your Formatted comments would look like with the current setup." & _
                             vbNewLine & vbNewLine & "Press 'Clear', enter your comments here, then press the 'Format!' button."
    cmdFormat_Click

End Sub

Private Sub Form_Resize()

'   The dark Frame at the top of the screen.
    With picTopBotFrame
        .Move 0, 0, Me.ScaleWidth, txtTopBotFormat.Height + picTopBotBars.Height + txtSideFormat.Height + 40
'       The Format bar at the top of the screen w/Blue text.
        txtTopBotFormat.Move 10, 10, .Width - 20
    End With  'picTopBotFormat

'   The Picture Frame containing the number of Top and Bottom Border Bars.
    With picTopBotBars
'       The Label with the number of Top and Bottom Border lines.
        lblTopBotBars.Move 5, 9
'       The Picture Frame containing the number of Top and Bottom Border Bars.
        .Move txtTopBotFormat.Left, txtTopBotFormat.Top + txtTopBotFormat.Height + 10, ((Me.ScaleWidth - lblTopBotFormat.Width) \ 2) - 40
'       The Text Box with the number of Top and Bottom Border lines.
        txtTopBotBars.Move .Width - 60, 5, 50, lblTopBotBars.Height
    End With  'picTopBotBars

    With picNumSpaces
'       The Picture Frame containing the number of spaces.
        .Move txtTopBotFormat.Left + txtTopBotFormat.Width - picTopBotBars.Width, picTopBotBars.Top, picTopBotBars.Width, picTopBotBars.Height
'       The Label with the number of Spaces.
        lblNumSpaces.Move 5, 3
'       The Text Box with the number of Spaces
        txtNumSpaces.Move .Width - 60, 5, 50, lblNumSpaces.Height
    End With  'picNumSpaces

    If lngScreenWide < 1100 Then
        With lblTopBotFormat
            .FontBold = False
            .FontSize = 10
            .Height = 17
            .Width = 320
        End With             'lblTopBotFormat

        With lblSideFormat
            .FontBold = False
            .FontSize = 10
            .Height = 17
            .Width = 320
        End With             'lblSideFormat
    End If

'   The info label line near the top of the screen.
    lblTopBotFormat.Move (Me.ScaleWidth - lblTopBotFormat.Width) \ 2, picTopBotBars.Top
'   The Label w/the format for the sides of the text box.
    lblSideFormat.Move lblTopBotFormat.Left, picTopBotBars.Top + picTopBotBars.Height - lblTopBotFormat.Height

'   The Text Box w/the format for the sides of the text box.
    txtSideFormat.Move txtTopBotFormat.Left, picTopBotBars.Top + picTopBotBars.Height + 10, txtTopBotFormat.Width
'   The Picture Frame containing the Command Boxes.
    picCommandBar.Move txtTopBotFormat.Left, picTopBotFrame.Top + picTopBotFrame.Height, txtTopBotFormat.Width

'   The "Clear" Key
    With cmdClear
        .Move 0, 0, picCommandBar.Width \ 5
'       The "Format" Key
        cmdFormat.Move .Left + .Width, 0, .Width
'       The "Copy" Key
        cmdCopy.Move cmdFormat.Left + .Width, 0, .Width
'       The "Save Text" Key
        cmdSaveText.Move cmdCopy.Left + .Width, 0, .Width
'       The "Restore Text" Key
        cmdRestoreText.Move cmdSaveText.Left + .Width, 0, .Width, 58
'   The Picture Frame containing the Rich Text Box.
    End With  'cmdClear

'   The Picture Frame containing the Rich Text Box.
    picEnterComments.Move picTopBotFrame.Left, picCommandBar.Top + picCommandBar.Height, picTopBotFrame.Width, Me.ScaleHeight - picCommandBar.Top - picCommandBar.Height
    chkLeftAlign.Move cmdRestoreText.Left + (cmdRestoreText.Width \ 2) - 132, 2
    lblEnterComments.Move lblTopBotFormat.Left, 5, lblTopBotFormat.Width, lblTopBotFormat.Height
'   The Rich Text Box where the input & output are displayed
    txtEnterComments.Move txtTopBotFormat.Left, lblEnterComments.Top + lblEnterComments.Height + 5, txtTopBotFormat.Width, picEnterComments.Height - txtEnterComments.Top - 10

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Close
    End

End Sub

Private Sub GetFormatBarInfo()

'   Determine the proper spacing and line length
    lngSpaces = txtNumSpaces.Text           'The number of spaces between the Border and the Text.
'   The last Space character before the Right Border.
    lngFieldEnd = InStrRev(strSideFormatBar, " ", -1, vbTextCompare)
'   The rightmost character of the Left Border.
    lngBorderEnd = InStr(1, strSideFormatBar, Space$(20), vbTextCompare) - 1
'   The Width of the Asterisk Border.
    lngLeftBorder = lngBorderEnd - lngSpaces

End Sub

Private Sub lblEnterComments_Click()

'   Move the Focus to the Text Box
    txtEnterComments.SetFocus

End Sub

Private Sub lblNumSpaces_Click()

    txtNumSpaces.SetFocus

End Sub

Private Sub lblTopBotBars_Click()

    txtTopBotBars.SetFocus

End Sub

Private Sub mnuAbout_Click()

    RG_About.Show vbModal, Me

End Sub

Private Sub mnuEditCopy_Click()

    Clipboard.Clear
    Clipboard.SetText txtEnterComments.SelText

End Sub

Private Sub mnuEditCut_Click()

    Clipboard.SetText txtEnterComments.SelText
    txtEnterComments.SelText = vbNullString

End Sub

Private Sub mnuEditPaste_Click()

    If Clipboard.GetFormat(vbCFText) Then
        txtEnterComments.SelText = Clipboard.GetText
    Else
        MsgBox "Invalid Clipboard format."
    End If

End Sub

Private Sub mnuEditSelectAll_Click()

    With txtEnterComments
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub mnuFileNew_Click()

    txtEnterComments.Text = vbNullString
    strOpenFile = vbNullString

End Sub

Private Sub mnuFileOpen_Click()

Dim lngFN1                   As Long

    On Error GoTo FileError
    With CDB
        .CancelError = True
        .Flags = cdlOFNFileMustExist
        .DefaultExt = "TXT"
        .Filter = "Text files|*.TXT|All files|*.*"
        .ShowOpen
        lngFN1 = FreeFile
        Open .FileName For Input As #lngFN1
    End With

    txtEnterComments.Text = Input$(LOF(lngFN1), #lngFN1)
    Close #lngFN1
    strOpenFile = CDB.FileName

Exit Sub

FileError:
    If Not Err.Number = cdlCancel Then
        MsgBox "Unkown error while opening file " & CDB.FileName
        strOpenFile = vbNullString
    End If

End Sub

Private Sub mnuFileSave_Click()

Dim lngFN1                   As Long

    If LenB(strOpenFile) = 0 Then
        mnuFileSaveAs_Click
    Else
        On Error GoTo FileError
        lngFN1 = FreeFile
        Open strOpenFile For Output As #lngFN1
        Print #lngFN1, txtEnterComments.Text
        Close #lngFN1

        GoTo Leave

FileError:
        If Err.Number <> cdlCancel Then
            MsgBox "Unkown error while saving file " & strOpenFile
            strOpenFile = vbNullString
        End If
    End If

Leave:

End Sub

Private Sub mnuFileSaveAs_Click()

Dim lngFN1                   As Long

    On Error GoTo FileError
    With CDB
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt
        .DefaultExt = "TXT"
        .Filter = "Text files|*.TXT|All files|*.*"
        .ShowSave

        lngFN1 = FreeFile
        Open .FileName For Output As #lngFN1
    End With

    Print #lngFN1, txtEnterComments.Text
    Close #lngFN1

    strOpenFile = CDB.FileName

Exit Sub

FileError:
    If Not Err.Number = cdlCancel Then

        MsgBox "Unkown error while saving file " & CDB.FileName
        strOpenFile = vbNullString

    End If

End Sub

Private Sub mnuHelp_Click()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Display the 'Help' page.                                     ****
'   ****                                                                 ****
'   ****    This is from a suggestion by:  Rob C.                        ****
'   ****                                                                 ****
'   ****    Thank you Rob C.                                             ****
'   ****                                                                 ****
'   *************************************************************************
'
    ShellExecute Me.hwnd, "open", App.Path & "\Formatter.Html", vbNullString, vbNullString, SW_MAXIMIZE

End Sub

Private Sub mnuQuit_Click()

    Unload Me
    End

End Sub

Private Sub mnuSample_Click()

Dim lngFN1                   As Long

    lngFN1 = FreeFile
    lngCounter = 0

    Open App.Path & "\Format.ter" For Input As #lngFN1
    txtEnterComments.Text = Input$(LOF(lngFN1), #lngFN1)
    Close #lngFN1

End Sub

Private Sub mnuUnFormat_Click()

Dim blnCheck    As Boolean
Dim lngFN1      As Long
Dim lngInStr    As Long
Dim lngLoop     As Long
Dim strArray()  As String
Dim strUnFormat As String

    lngFN1 = FreeFile
    Open App.Path & "\T.emp" For Output As #lngFN1
        Print #lngFN1, txtEnterComments.Text
    Close #lngFN1

    txtEnterComments.Text = vbNullString
    lngFN1 = FreeFile
    Open App.Path & "\T.emp" For Input As #lngFN1

    Do Until EOF(lngFN1)
        Line Input #lngFN1, strUnFormat
'       Remove any spaces from the front and back of the line.
        strUnFormat = Trim$(strUnFormat)
'       Look for the Tick Mark (') at the front of the line.
        lngInStr = InStr(1, strUnFormat, "'", vbTextCompare)

        If lngInStr < 1 Then
            GoTo Read_Next
        End If

        If Len(strUnFormat) > 0 Then
'           Remove the Tick Mark (') and any spaces from the front of the line.
            strUnFormat = Trim$(Mid$(strUnFormat, lngInStr + 1, Len(strUnFormat)))
        End If

'       This will handle a line that has 1 or more spaces before the "'".
        If Len(strUnFormat) = 0 Then
            GoTo Read_Next
        End If

'       Find the Asterisks (*) in the line.
        strArray = Split(strUnFormat, "*", , vbTextCompare)

        For lngLoop = 1 To UBound(strArray)
            strUnFormat = Trim(strArray(lngLoop))
'           Pull out only the stripped line.
            If Len(strUnFormat) > 0 Then
                Exit For
            End If
        Next lngLoop

        If blnCheck = False Then
            If Len(strUnFormat) = 0 Then
                GoTo Read_Next
            End If

            If Asc(Left$(strUnFormat, 1)) < 14 Then
                GoTo Read_Next
            End If
        End If

        txtEnterComments.Text = txtEnterComments.Text & strUnFormat & vbNewLine
        blnCheck = True
        ReDim strArray(0)

Read_Next:
    Loop

    Close #lngFN1
    DeleteFile App.Path & "\T.emp"
    txtEnterComments.Text = Trim$(txtEnterComments.Text)

End Sub

Private Sub PadLine()

Dim strComments              As String
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is where I insert all the asterisks "*" and spacing.    ****
'   ****    In other words, this is where I format your comments.        ****
'   ****                                                                 ****
'   *************************************************************************
'
    If strBreakerBar = "Place a BreakerBar here!" Then
        strComments = strTopFormatBar
        GoTo PrintLine
    End If

'   "strComments" is the comment broken at its last space character.
'   Then set to its formatted length.
    strComments = Left$(Left$(strSideFormatBar, lngBorderEnd + lngSpaces) & Left$(LTrim$(Left$(strBreakerBar, lngLoop2 - 1)) & Space$(lngFieldSize), lngFieldSize) & Space$(99), lngFieldEnd) & Right$(strSideFormatBar, Len(strSideFormatBar) - lngFieldEnd)

PrintLine:
'   This is the comment line w/all necessary padding
    txtEnterComments.Text = txtEnterComments.Text & vbNewLine & strComments

End Sub

Private Sub picNumSpaces_Click()

    txtNumSpaces.SetFocus

End Sub

Private Sub picTopBotBars_Click()

    txtTopBotBars.SetFocus

End Sub

Private Sub txtNumSpaces_Change()

'   Determine the proper spacing and line length
    lngSpaces = txtNumSpaces.Text           'The number of spaces between the Border and the Text.

End Sub

Private Sub txtSideFormat_Change()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Make all of the Format Bars the same length.                 ****
'   ****                                                                 ****
'   *************************************************************************
'
    With txtSideFormat
'       This will limit the length of the Format Bars (currently500)
        If Len(.Text) > 500 Then
            .Text = Trim$(Left$(strSideFormatBar, 500))
        Else
            strSideFormatBar = .Text
        End If

'       Get some basic Format Bar Info.
        GetFormatBarInfo

        If Len(.Text) <> Len(txtTopBotFormat.Text) Then
            txtTopBotFormat.Text = Left$(Left$(.Text, InStr(1, .Text, Space$(20), vbTextCompare) - 1) & String$(Len(.Text), "*"), Len(.Text))
        End If
    End With  'txtSideFormat

End Sub

Private Sub txtTopBotFormat_Change()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Make all of the Format Bars the same length.                 ****
'   ****                                                                 ****
'   *************************************************************************
'
    With txtTopBotFormat
'       This will limit the length of the Format Bars (currently 500)
        If Len(.Text) > 500 Then
            .Text = Trim$(Left$(strTopFormatBar, 500))
        Else
            strTopFormatBar = .Text
        End If

        If Len(.Text) <> Len(txtSideFormat.Text) Then
            txtSideFormat.Text = Left$("'" & Space$(lngSpaces - 1) & String$(lngLeftBorder, "*") & Space$(lngFieldEnd - lngBorderEnd) & String$(lngFieldEnd, "*"), Len(strTopFormatBar))
        End If
    End With  'txtTopBotFormat

End Sub
