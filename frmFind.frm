VERSION 5.00
Begin VB.Form frmFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7080
   Icon            =   "frmFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTooltips 
      Caption         =   "Tooltips"
      Height          =   285
      Left            =   5505
      TabIndex        =   39
      Top             =   3405
      Value           =   2  'Grayed
      Width           =   1440
   End
   Begin VB.Frame fraAllframes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   150
      TabIndex        =   14
      Top             =   540
      Width           =   5280
      Begin VB.CommandButton btnAdvancedFeatures 
         Caption         =   "v"
         Height          =   285
         Left            =   4695
         TabIndex        =   40
         Top             =   15
         Width           =   360
      End
      Begin VB.ComboBox cmbDirection 
         Height          =   315
         Left            =   3555
         TabIndex        =   35
         Text            =   "All"
         Top             =   15
         Width           =   1095
      End
      Begin VB.Frame fraOrigin 
         Caption         =   "Origin"
         Height          =   1110
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Width           =   2415
         Begin VB.OptionButton optFromCursor 
            Caption         =   "From Cursor Pos&ition"
            Height          =   285
            Left            =   270
            TabIndex        =   34
            Top             =   315
            Width           =   1905
         End
         Begin VB.OptionButton optFromTop 
            Caption         =   "From th&e Top"
            Height          =   270
            Left            =   270
            TabIndex        =   33
            Top             =   630
            Value           =   -1  'True
            Width           =   1905
         End
      End
      Begin VB.Frame fraScope 
         Caption         =   "Scope"
         Height          =   2115
         Left            =   0
         TabIndex        =   26
         Top             =   1605
         Width           =   2415
         Begin VB.OptionButton optScope 
            Caption         =   "Current &Procedure"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   31
            Top             =   375
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.OptionButton optScope 
            Caption         =   "Current &Module"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   30
            Top             =   705
            Width           =   1995
         End
         Begin VB.OptionButton optScope 
            Caption         =   "&Current Project"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   29
            Top             =   1020
            Width           =   1995
         End
         Begin VB.OptionButton optScope 
            Caption         =   "Fol&der"
            Height          =   195
            Index           =   5
            Left            =   270
            TabIndex        =   28
            Top             =   1650
            Width           =   1995
         End
         Begin VB.OptionButton optScope 
            Caption         =   "Selected &Text"
            Enabled         =   0   'False
            Height          =   195
            Index           =   4
            Left            =   270
            TabIndex        =   27
            Top             =   1335
            Width           =   1995
         End
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   2040
         Left            =   2580
         TabIndex        =   19
         Top             =   360
         Width           =   2550
         Begin VB.CheckBox chkCaseSensitive 
            Caption         =   "Match Ca&se"
            Height          =   255
            Left            =   255
            TabIndex        =   25
            Top             =   270
            Width           =   2145
         End
         Begin VB.CheckBox chkWholeWords 
            Caption         =   "Find Whole Words &Only"
            Height          =   300
            Left            =   255
            TabIndex        =   24
            Top             =   525
            Width           =   2220
         End
         Begin VB.CheckBox chkSkipComments 
            Caption         =   "Skip Comments"
            Height          =   255
            Left            =   255
            TabIndex        =   23
            Top             =   840
            Width           =   2145
         End
         Begin VB.CheckBox chkSkipTags 
            Caption         =   "Skip Tags"
            Height          =   300
            Left            =   255
            TabIndex        =   22
            Top             =   1095
            Width           =   2220
         End
         Begin VB.CheckBox chkPatternMatching 
            Caption         =   "&Use Pattern Matching"
            Height          =   300
            Left            =   255
            TabIndex        =   21
            Top             =   1650
            Width           =   2220
         End
         Begin VB.CheckBox chkSkipStrings 
            Caption         =   "Skip Strings"
            Height          =   300
            Left            =   255
            TabIndex        =   20
            Top             =   1380
            Width           =   2220
         End
      End
      Begin VB.Frame fraOutput 
         Caption         =   "Output"
         Height          =   1275
         Left            =   2580
         TabIndex        =   15
         Top             =   2445
         Width           =   2565
         Begin VB.OptionButton optListAll 
            Caption         =   "List All items found"
            Height          =   195
            Left            =   225
            TabIndex        =   18
            Top             =   315
            Width           =   1995
         End
         Begin VB.OptionButton optHighlightFound 
            Caption         =   "Highlight all found"
            Height          =   195
            Left            =   225
            TabIndex        =   17
            Top             =   615
            Width           =   1995
         End
         Begin VB.OptionButton optFindNext 
            Caption         =   "Find Next"
            Height          =   195
            Left            =   225
            TabIndex        =   16
            Top             =   915
            Value           =   -1  'True
            Width           =   1995
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Direction:"
         Height          =   345
         Left            =   2625
         TabIndex        =   36
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbReplace 
      Height          =   315
      Left            =   1260
      TabIndex        =   37
      Text            =   "something totally diffferent"
      Top             =   720
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.CommandButton btnReplaceAll 
      Caption         =   "Replace &All"
      Enabled         =   0   'False
      Height          =   420
      Left            =   5430
      TabIndex        =   13
      Top             =   1845
      Width           =   1530
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "&Help"
      Height          =   420
      Left            =   5430
      TabIndex        =   12
      Top             =   2490
      Width           =   1545
   End
   Begin VB.Frame fraFolderWildcard 
      Caption         =   "Folder and Wildcard Options"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   150
      TabIndex        =   6
      Top             =   4455
      Width           =   6780
      Begin VB.CheckBox chkSubFolders 
         Caption         =   "Include Sub-Folders"
         Enabled         =   0   'False
         Height          =   300
         Left            =   330
         TabIndex        =   11
         Top             =   780
         Width           =   2220
      End
      Begin VB.ComboBox cmbWildcard 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5055
         TabIndex        =   10
         Text            =   "*.*"
         Top             =   360
         Width           =   1545
      End
      Begin VB.CommandButton btnFolder 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   360
         Left            =   4500
         TabIndex        =   9
         Top             =   330
         Width           =   360
      End
      Begin VB.ComboBox cmbFolder 
         Enabled         =   0   'False
         Height          =   315
         Left            =   795
         TabIndex        =   8
         Text            =   "c:\vb6\exclusions"
         Top             =   360
         Width           =   3645
      End
      Begin VB.Label lblFolder 
         Caption         =   "Folder:"
         Enabled         =   0   'False
         Height          =   345
         Left            =   180
         TabIndex        =   7
         Top             =   375
         Width           =   1275
      End
   End
   Begin VB.CommandButton btnReplace 
      Caption         =   "&Replace"
      Height          =   420
      Left            =   5460
      TabIndex        =   5
      Top             =   1335
      Width           =   1500
   End
   Begin VB.CommandButton btnFindMenu 
      Caption         =   ">"
      Height          =   330
      Left            =   4875
      TabIndex        =   4
      Top             =   120
      Width           =   330
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   420
      Left            =   5445
      TabIndex        =   3
      Top             =   585
      Width           =   1500
   End
   Begin VB.CommandButton btnFind 
      Caption         =   "Fi&nd"
      Height          =   420
      Left            =   5430
      TabIndex        =   2
      Top             =   75
      Width           =   1500
   End
   Begin VB.ComboBox cmbSearchTerm 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Text            =   "function(a)"
      Top             =   150
      Width           =   3585
   End
   Begin VB.Label lblReplaceWith 
      Caption         =   "Replace With:"
      Height          =   345
      Left            =   165
      TabIndex        =   38
      Top             =   765
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblFindWhat 
      Caption         =   "Find What:"
      Height          =   345
      Left            =   165
      TabIndex        =   1
      Top             =   195
      Width           =   1275
   End
   Begin VB.Menu mnuTopMenu 
      Caption         =   "mnuTopMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuPinItem 
         Caption         =   "Pin item to list"
      End
      Begin VB.Menu mnuUnPin 
         Caption         =   "Un-Pin Item"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete current item from list"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear Item"
      End
   End
   Begin VB.Menu mnuPrefsmenu 
      Caption         =   "mnuPrefsmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdvancedTop 
         Caption         =   "Advanced Features "
         Begin VB.Menu mnuAdvancedON 
            Caption         =   "ON"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuAdvancedOFF 
            Caption         =   "OFF"
         End
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form      : frmFind
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public searchType As String
Public btnReplaceClicked As Boolean

' key presses
Public CTRL_1 As Boolean


'---------------------------------------------------------------------------------------
' Procedure : btnAdvancedFeatures_Click
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAdvancedFeatures_Click()
   On Error GoTo btnAdvancedFeatures_Click_Error

    If mnuAdvancedON.Checked = True Then
        Call mnuAdvancedOFF_Click
    Else
        Call mnuAdvancedON_Click
    End If

   On Error GoTo 0
   Exit Sub

btnAdvancedFeatures_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAdvancedFeatures_Click of Form frmFind"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnClose_Click
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_Click()
   On Error GoTo btnClose_Click_Error

    Unload frmFind
    Set frmFind = Nothing

   On Error GoTo 0
   Exit Sub

btnClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClose_Click of Form frmFind"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnFindMenu_Click
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnFindMenu_Click()
   On Error GoTo btnFindMenu_Click_Error

    Me.PopupMenu mnuTopMenu, vbPopupMenuRightButton

   On Error GoTo 0
   Exit Sub

btnFindMenu_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFindMenu_Click of Form frmFind"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnReplace_Click
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnReplace_Click()

   On Error GoTo btnReplace_Click_Error

    btnReplaceClicked = True
    
    If mnuAdvancedON.Checked = True Then
        Call makeElementsAvailable("advanced")
    Else
        Call makeElementsAvailable("simple")
    End If
    
    frmFind.Refresh

   On Error GoTo 0
   Exit Sub

btnReplace_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnReplace_Click of Form frmFind"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkTooltips_Click
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkTooltips_Click()
   On Error GoTo chkTooltips_Click_Error

    Call setTooltips

   On Error GoTo 0
   Exit Sub

chkTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkTooltips_Click of Form frmFind"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_KeyDown
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo Form_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

   On Error GoTo 0
   Exit Sub

Form_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_KeyDown of Form frmFind"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    searchType = "advanced"
    
    ' key presses
    CTRL_1 = False
    
    btnReplaceClicked = False
    frmFind.Top = 705
    chkTooltips.Value = 1
    cmbSearchTerm.AddItem "Sample search term 1", 0
    cmbSearchTerm.AddItem "Sample search term 2", 1
    cmbSearchTerm.AddItem "Sample search term 3", 2
    cmbReplace.AddItem "Sample replacement text 1", 0
    cmbReplace.AddItem "Sample replacement text 2", 1
    cmbReplace.AddItem "Sample replacement text 3", 2
    cmbDirection.AddItem "All", 0
    cmbDirection.AddItem "Up", 1
    cmbDirection.AddItem "Down", 2
    
    If optFindNext.Value = True Then
        btnFind.Caption = "Find &Next"
    Else
        btnFind.Caption = "Find &All"
    End If
    
    Call makeElementsAvailable(searchType)
    
    Call setTooltips

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form frmFind"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : makeElementsAvailable
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub makeElementsAvailable(ByVal thisType As String)
   On Error GoTo makeElementsAvailable_Error

    frmFind.Visible = False
    If thisType = "simple" Then
        frmFind.Caption = "Find"
        btnAdvancedFeatures.Caption = "v"
        
        fraAllframes.Top = 540
        frmFind.Height = 2850
        
        fraOrigin.Visible = False
        fraOutput.Visible = False
        optScope(5).Visible = False
        chkSkipComments.Visible = False
        chkSkipTags.Visible = False
        chkSkipStrings.Visible = False
        btnReplaceAll.Visible = False
        btnFindMenu.Visible = False
        fraFolderWildcard.Visible = False
        chkTooltips.Visible = False
        
        fraScope.Top = 0
        chkPatternMatching.Top = 840
        btnHelp.Top = 1800
        
        fraScope.Height = 1700
        fraOptions.Height = 1310
        
        If btnReplaceClicked = True Then
        
            frmFind.Caption = "Replace"
            btnReplaceAll.Enabled = True
            fraFolderWildcard.Top = 5145
            
            lblReplaceWith.Visible = True
            cmbReplace.Visible = True

            btnReplaceAll.Visible = True
            frmFind.Height = 2850 + 695
            fraAllframes.Top = 540 + 695
            btnHelp.Top = 2400
        End If
    Else
        frmFind.Caption = "Find"
        btnAdvancedFeatures.Caption = "^"
        
        frmFind.Height = 6255
        fraAllframes.Top = 540
        fraScope.Top = 1605
        chkPatternMatching.Top = 1650
        btnHelp.Top = 2490
        fraFolderWildcard.Top = 4455
        
        fraScope.Height = 2115
        fraOptions.Height = 2040
        fraOrigin.Visible = True
        fraOutput.Visible = True
        optScope(5).Visible = True
        chkSkipComments.Visible = True
        chkSkipTags.Visible = True
        chkSkipStrings.Visible = True
        btnReplaceAll.Visible = True
        btnFindMenu.Visible = True
        fraFolderWildcard.Visible = True
        chkTooltips.Visible = True
        
        If btnReplaceClicked = True Then
        
            frmFind.Caption = "Replace"
            btnReplaceAll.Enabled = True
            fraFolderWildcard.Top = 5145
            
            lblReplaceWith.Visible = True
            cmbReplace.Visible = True
            fraAllframes.Top = 540 + 695
            frmFind.Height = 6255 + 695

            chkPatternMatching.Top = 1650
            btnHelp.Top = btnHelp.Top + 300
        End If
    End If
    frmFind.Visible = True
    frmFind.Refresh
    btnFind.SetFocus ' to remove the focus from the basic/advanced button
    frmFind.SetFocus ' to allow the form to respond to key presses

   On Error GoTo 0
   Exit Sub

makeElementsAvailable_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeElementsAvailable of Form frmFind"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setTooltips
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setTooltips()
   On Error GoTo setTooltips_Error

    If chkTooltips.Value = 1 Then
        btnFind.ToolTipText = "Click this button to commence the search"
        btnClose.ToolTipText = "This will close this search form"
        btnReplace.ToolTipText = "This will replace only the next occurrence of the search string with the newly supplied text"
        btnReplaceAll.ToolTipText = "This will replace all found occurrences of the search string - instantly"
        btnHelp.ToolTipText = "Click for help"
        chkCaseSensitive.ToolTipText = "This option makes the search case sensitive"
        chkWholeWords.ToolTipText = "This option alters the search to show only whole words"
        chkSkipComments.ToolTipText = "This option alters the search to skip comments completely"
        chkSkipTags.ToolTipText = "This option alters the search to skip tags completely"
        chkSkipStrings.ToolTipText = "This option alters the search to skip all strings completely"
        chkPatternMatching.ToolTipText = "Pattern-matching allows you to match each character against a specific character."
        optListAll.ToolTipText = "This option causes the search list to be populated so that you can see all matching results"
        optHighlightFound.ToolTipText = "This option merely highlights all matching results in the code"
        optFindNext.ToolTipText = "This is the default search behaviour, just searching for one match at a time, VB6 style"
        btnFindMenu.ToolTipText = "The Find Menu allowing you to edit the search list"
        cmbDirection.ToolTipText = "Select the search direction"
        cmbSearchTerm.ToolTipText = "Enter the Text you want to find in your code"
        cmbReplace.ToolTipText = "Enter the text that will replace your search term"
        optFromCursor.ToolTipText = "Search from the cursor position"
        optFromTop.ToolTipText = "Search from the top"
        optScope(1).ToolTipText = "Search only within the current subroutine or function"
        optScope(2).ToolTipText = "Search only within the current module or class"
        optScope(3).ToolTipText = "Search within the whole project"
        optScope(4).ToolTipText = "Search only within the user-selected text"
        optScope(5).ToolTipText = "Extend the search to a folder, selected in the section below, enabled when this option is selected "
        cmbFolder.ToolTipText = "This box shows any currently selected folder to search"
        btnFolder.ToolTipText = "Select a folder to search"
        cmbWildcard.ToolTipText = "Set a wildcard (eg. *.*) to select matching files to search"
        chkSubFolders.ToolTipText = "If you wish to search all sub-folders, click here"
        btnAdvancedFeatures.ToolTipText = "Press to toggle between advanced or basic search features"
    Else
        btnFind.ToolTipText = vbNullString
        btnClose.ToolTipText = vbNullString
        btnReplace.ToolTipText = vbNullString
        btnReplaceAll.ToolTipText = vbNullString
        btnHelp.ToolTipText = vbNullString
        chkCaseSensitive.ToolTipText = vbNullString
        chkWholeWords.ToolTipText = vbNullString
        chkSkipComments.ToolTipText = vbNullString
        chkSkipTags.ToolTipText = vbNullString
        chkSkipStrings.ToolTipText = vbNullString
        chkPatternMatching.ToolTipText = vbNullString
        optListAll.ToolTipText = vbNullString
        optHighlightFound.ToolTipText = vbNullString
        optFindNext.ToolTipText = vbNullString
        btnFindMenu.ToolTipText = vbNullString
        cmbDirection.ToolTipText = vbNullString
        cmbSearchTerm.ToolTipText = vbNullString
        cmbReplace.ToolTipText = vbNullString
        optFromCursor.ToolTipText = vbNullString
        optFromTop.ToolTipText = vbNullString
        optScope(1).ToolTipText = vbNullString
        optScope(2).ToolTipText = vbNullString
        optScope(3).ToolTipText = vbNullString
        optScope(4).ToolTipText = vbNullString
        optScope(5).ToolTipText = vbNullString
        cmbFolder.ToolTipText = vbNullString
        btnFolder.ToolTipText = vbNullString
        cmbWildcard.ToolTipText = vbNullString
        chkSubFolders.ToolTipText = vbNullString
        btnAdvancedFeatures.ToolTipText = vbNullString
    End If

   On Error GoTo 0
   Exit Sub

setTooltips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setTooltips of Form frmFind"
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuPrefsmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraAllframes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuPrefsmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuPrefsmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraOrigin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuPrefsmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraOutput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuPrefsmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraScope_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuPrefsmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAdvancedON_Click
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAdvancedON_Click()
   On Error GoTo mnuAdvancedON_Click_Error

    searchType = "advanced"
    mnuAdvancedOFF.Checked = False
    mnuAdvancedON.Checked = True
    Call makeElementsAvailable(searchType)

   On Error GoTo 0
   Exit Sub

mnuAdvancedON_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAdvancedON_Click of Form frmFind"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAdvancedOFF_Click
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAdvancedOFF_Click()
   On Error GoTo mnuAdvancedOFF_Click_Error

    searchType = "simple"
    mnuAdvancedOFF.Checked = True
    mnuAdvancedON.Checked = False
    Call makeElementsAvailable(searchType)

   On Error GoTo 0
   Exit Sub

mnuAdvancedOFF_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAdvancedOFF_Click of Form frmFind"
End Sub

Private Sub optFindNext_Click()
        btnFind.Caption = "Find &Next"
End Sub

Private Sub optHighlightFound_Click()
        btnFind.Caption = "&Find All"
End Sub

Private Sub optListAll_Click()
        btnFind.Caption = "&Find All"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optScope_Click
' Author    : beededea
' Date      : 02/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optScope_Click(Index As Integer)
   On Error GoTo optScope_Click_Error

    If optScope(5).Value = True Then
        fraFolderWildcard.Enabled = True
        cmbFolder.Enabled = True
        btnFolder.Enabled = True
        cmbWildcard.Enabled = True
        chkSubFolders.Enabled = True
        lblFolder.Enabled = True
    Else
        fraFolderWildcard.Enabled = False
        cmbFolder.Enabled = False
        btnFolder.Enabled = False
        cmbWildcard.Enabled = False
        chkSubFolders.Enabled = False
        lblFolder.Enabled = False
    End If

   On Error GoTo 0
   Exit Sub

optScope_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optScope_Click of Form frmFind"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : getkeypress
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : getting a keypress from the keyboard
'---------------------------------------------------------------------------------------
'
Public Sub getKeyPress(ByVal KeyCode As Integer, ByVal Shift As Integer)
    
    On Error GoTo getkeypress_Error

    Select Case KeyCode
        Case vbKeyControl
            CTRL_1 = True

        Case vbKeyF ' ctrl+F for find
            If CTRL_1 = True Then
                btnReplaceClicked = False
                Call makeElementsAvailable(searchType)
                CTRL_1 = False
            Else
                ' folder
            End If
        Case vbKeyH ' ctrl+H for replace
            If CTRL_1 = True Then
                btnReplaceClicked = True
                Call makeElementsAvailable(searchType)
                CTRL_1 = False
            End If
        Case vbKeyR ' ctrl+R for replace
            If CTRL_1 = True Then
                btnReplaceClicked = True
                Call makeElementsAvailable(searchType)
                CTRL_1 = False
            End If
    End Select
    
 
    On Error GoTo 0
   Exit Sub

getkeypress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getkeypress of Module module1"

End Sub
