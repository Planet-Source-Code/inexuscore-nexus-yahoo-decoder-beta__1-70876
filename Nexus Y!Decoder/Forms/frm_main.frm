VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{AD511FF1-C0E0-4DA3-899A-80C2675DCE9A}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frm_main 
   Caption         =   "Nexus Yahoo! Decoder (Beta)"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin HookMenu.ctxHookMenu HookMenu1 
      Left            =   1320
      Top             =   4680
      _ExtentX        =   900
      _ExtentY        =   900
      MenuGradientColor=   0
      MenuForeColor   =   -2147483640
      MenuBorderColor =   -2147483632
      MenuGradientSelectColor=   0
      PopupBorderColor=   -2147483632
      PopupBorderColor=   -2147483640
      PopupGradientSelectColor=   0
      SideBarColor    =   14215660
      SideBarGradientColor=   0
      CheckForeColor  =   -2147483641
      ShadowColor     =   0
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList i24x24 
      Left            =   3180
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":7D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":96D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":9E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":A5C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":AD42
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":B4BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":BC36
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":C3B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":DD42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":F6D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1173E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":137A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":13F22
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbMain 
      Align           =   1  'Align Top
      DragMode        =   1  'Automatic
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   1720
      BandCount       =   2
      _CBWidth        =   10140
      _CBHeight       =   975
      _Version        =   "6.0.8169"
      Child1          =   "tbrMain"
      MinHeight1      =   450
      Width1          =   3135
      NewRow1         =   0   'False
      Child2          =   "picPath"
      MinHeight2      =   435
      Width2          =   3975
      NewRow2         =   -1  'True
      Begin VB.PictureBox picPath 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   165
         ScaleHeight     =   435
         ScaleWidth      =   9885
         TabIndex        =   5
         Top             =   510
         Width           =   9885
         Begin VB.TextBox txtPath 
            Height          =   315
            Left            =   540
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   60
            Width           =   4635
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse"
            Height          =   315
            Left            =   5220
            TabIndex        =   6
            Top             =   60
            Width           =   1035
         End
         Begin VB.Label lblPath 
            AutoSize        =   -1  'True
            Caption         =   "Path :"
            Height          =   195
            Left            =   60
            TabIndex        =   8
            Top             =   120
            Width           =   435
         End
      End
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   450
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   794
         ButtonWidth     =   820
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "i24x24"
         HotImageList    =   "i24x24"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save Current Archive"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Export"
               Object.ToolTipText     =   "Export Decoded Archive"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ExportHTML"
                     Text            =   "Export as HTML"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ExportText"
                     Text            =   "Export as Text"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ExportRTF"
                     Text            =   "Export as RTF"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "CopySelection"
               Object.ToolTipText     =   "Copy Selection"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Refresh"
               Object.ToolTipText     =   "Refresh"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindText"
               Object.ToolTipText     =   "Find Text"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Backup"
               Object.ToolTipText     =   "Backup Archive Data"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ProfileEnabler"
               Object.ToolTipText     =   "Enable Profiles Archiving"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "About"
               Object.ToolTipText     =   "About Application"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Exit"
               Object.ToolTipText     =   "Exit Application"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   2520
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1469C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":14C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":151D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1576A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":15D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1629E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":16838
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":16DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1716C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":17506
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   7065
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17357
            Key             =   "Status"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   1920
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbMessages 
      Height          =   2655
      Left            =   3780
      TabIndex        =   1
      Top             =   1200
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   4683
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDropMode     =   0
      TextRTF         =   $"frm_main.frx":178A0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwProfiles 
      Height          =   2595
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   4577
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "i16x16"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Archives"
      End
      Begin VB.Menu mnuEnable 
         Caption         =   "Enable Profiles"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SaveSettings()
    Dim X As Integer    '// free file handle
    X = FreeFile        '// create a new free file
    
    '// update screen mouse pointer
    Screen.MousePointer = vbHourglass
    '// open the app config file for output access
    Open App.Path & "\Config.ini" For Output As #1
        '// print Profiles Directory property and its value
        Print #1, "[Profiles Directory]"
        Print #1, "Path=" & Me.txtPath.Text
    Close #1    '// close config file
    '// update screen mouse pointer
    Screen.MousePointer = vbDefault
End Sub
Private Sub LoadSettings()
    '// if app config file doesnt exist
    If Dir(App.Path & "\Config.ini") = "" Then
        '// set the profiles directory path to default
        Me.txtPath.Text = "C:\Program Files\Yahoo!\Messenger\Profiles\"
        strFolder = Me.txtPath.Text '// update strFolder value
        Exit Sub    '// exit this procedure
    Else    '// if app config path exists
        Dim X As Integer    '// free file handle
        Dim strTemp As String   '// temp string
        Dim intIndex As Integer '// index var
        
        '// update screen mouse pointer
        Screen.MousePointer = vbHourglass
        '// create a new free file
        X = FreeFile
        '// open app config file for input access
        Open App.Path & "\Config.ini" For Input As #1
            Line Input #1, strTemp  '// read a line from input
            Line Input #1, strTemp  '// read a second line from input
            '// if gathered line is not a property tag "[]"
            If Not Left(strTemp, 1) = "[" And Not Right(strTemp, 1) = "]" Then
                '// get the starting index of value
                intIndex = InStr(1, strTemp, "=", vbBinaryCompare)
                '// if value exists
                If intIndex <> 0 Then
                    '// get the value string
                    strTemp = Mid(strTemp, intIndex + 1, Len(strTemp))
                    '// update root path value
                    Me.txtPath.Text = strTemp
                    strFolder = strTemp '// update strFolder value
                End If
            End If
        Close #1    '// close config file
        '// update screen mouse pointer
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Function DirSearch(ByVal strDir As String) As Collection
    '// update screen mouse pointer
    Screen.MousePointer = vbHourglass
    
    '//a dd all files and directories
    '// to the function's returning collection
    Set DirSearch = New Collection
    strDir = Dir(strDir, vbDirectory + vbNormal)
    
    Do Until strDir = ""
        If strDir <> "." And strDir <> ".." Then
            Call DirSearch.Add(strDir)
        End If
        
        strDir = Dir()
    Loop
    '// update screen mouse pointer
    Me.MousePointer = vbDefault
End Function

Private Sub LoadProfiles(ByVal strPath As String)
    '// recursively search through files/directories in "strPath"
    Dim cProf1 As New Collection    '// collection object
    Dim cProf2 As New Collection    '// collection object
    Dim cProf3 As New Collection    '// collection object
    Dim cItem1 As Variant           '// variant object
    Dim cItem2 As Variant           '// variant object
    Dim cItem3 As Variant           '// variant object
    
    '// update screen mouse pointer
    Screen.MousePointer = vbHourglass
    
    '// update main status bar text
    sbMain.Panels(1).Text = "Looking for available profiles..."
    '// clear treeview nodes
    tvwProfiles.Nodes.Clear
    '// populate the first collection object
    '// by searching through the root path
    Set cProf1 = DirSearch(strPath)
    '// loop through the first collection values
    For Each cItem1 In cProf1
        '// add the profile nodes
        Call tvwProfiles.Nodes.Add(, , cItem1, cItem1, 8, 8)
        '// populate the second collection object
        '// by searching for current profile's archive buddies
        Set cProf2 = DirSearch(strPath & cItem1 & "\Archive\Messages\")
        '// loop through the second collection values
        For Each cItem2 In cProf2
            '// add the buddy nodes
            Call tvwProfiles.Nodes.Add(cItem1, tvwChild, cItem1 & "," & cItem2, cItem2, 9, 9)
            '// populate the third collection object
            '// by searching through the current buddy's archive messages
            Set cProf3 = DirSearch(strPath & cItem1 & "\Archive\Messages\" & cItem2 & "\")
            '// loop through the third collection object
            For Each cItem3 In cProf3
                '// add the archive messages nodes
                Call tvwProfiles.Nodes.Add(cItem1 & "," & cItem2, tvwChild, _
                    cItem1 & "," & cItem2 & "," & cItem3, _
                    Left(cItem3, Len(cItem3) - (Len(cItem1) + 5)), 10, 10)
            Next cItem3
        Next cItem2
    Next cItem1
    
    '// dispose used variables and objects
    Set cItem1 = Nothing
    Set cItem2 = Nothing
    Set cItem3 = Nothing
    Set cProf1 = Nothing
    Set cProf2 = Nothing
    Set cProf3 = Nothing
    
    '// update main status bar text
    sbMain.Panels(1).Text = "Ready"
    '// update sceen mouse pointer
    Screen.MousePointer = vbDefault
End Sub
Private Sub ExportToHTML()
    On Error GoTo hell
    
    '// if data textbox is not empty
    If rtbMessages.Text <> "" Then
        '// get the selected treeview node data (profile, buddy, archive)
        strNodeData = Split(tvwProfiles.SelectedItem.Key, ",")
        '// if the selected node is a message node (archive files)
        If UBound(strNodeData) = 2 Then
            '// initialize CommonDialog props
            With Cdl
                .DialogTitle = "Save Decoded Archive As..."
                .CancelError = True
                .Flags = &H2
                .Filter = "HTML Documents(*.htm)|*.htm"
                .DefaultExt = "*.htm"
                .ShowSave
                '// if output filename is specified
                If .FileName <> "" Then
                    '// decode the selected archive file as HTML
                    Call DecodeAsHTML(strNodeData(0), strFolder & strNodeData(0) & "\Archive\Messages\" & strNodeData(1) & "\" & strNodeData(2), .FileName)
                    '// command completed, prompt user
                    MsgBox "YM Archive Saved Successfully", vbInformation + vbOKOnly, "Save Decoded Archive"
                End If
            End With
        End If
    Else    '// if data textbox is empty
        MsgBox "Please select an archive file to decode first", vbInformation + vbOKOnly, "Export To HTML"
    End If
'// error handler label
hell:
    If Err.Number = 0 Or Err.Number = 32755 Then
        Resume Next
    End If
End Sub
Private Sub ExportToText()
    On Error GoTo hell
    
    '// if data textbox is not empty
    If rtbMessages.Text <> "" Then
        '// get the selected node's data (profile, buddy, archive)
        strNodeData = Split(tvwProfiles.SelectedItem.Key, ",")
        '// if selected node is a message node (archive files)
        If UBound(strNodeData) = 2 Then
            '// initialize CommonDialog props
            With Cdl
                .DialogTitle = "Save Decoded Archive As..."
                .CancelError = True
                .Flags = &H2
                .Filter = "Text Documents(*.txt)|*.txt"
                .DefaultExt = "*.txt"
                .ShowSave
                '// if output filename is specified
                If .FileName <> "" Then
                    Dim X As Integer    '// free file handle
                    X = FreeFile    '// create a new free file
                    '// open output file for output access
                    Open .FileName For Output As #X
                        '// print the data textbox contents
                        '// as the output data to the file
                        Print #X, rtbMessages.Text
                    Close #X    '// close the output file
                    '// command completed, prompt user
                    MsgBox "YM Archive Saved Successfully", vbInformation + vbOKOnly, "Save Decoded Archive"
                End If
            End With
        End If
    Else    '// if data textbox is empty
        MsgBox "Please select an archive file to decode first", vbInformation + vbOKOnly, "Export To HTML"
    End If
'// error handler label
hell:
    If Err.Number = 0 Or Err.Number = 32755 Then
        Resume Next
    End If
End Sub
Private Sub ExportToRTF()
    On Error GoTo hell
    
    '// if data textbox is not empty
    If rtbMessages.Text <> "" Then
        '// get the selected node's data (profile, buddy, archive)
        strNodeData = Split(tvwProfiles.SelectedItem.Key, ",")
        '// if selected node is a message node (archive files)
        If UBound(strNodeData) = 2 Then
            '// initialize CommonDialog props
            With Cdl
                .DialogTitle = "Save Decoded Archive As..."
                .CancelError = True
                .Flags = &H2
                .Filter = "RTF Documents(*.rtf)|*.rtf"
                .DefaultExt = "*.htm"
                .ShowSave
                '// if output filename is specified
                If .FileName <> "" Then
                    '// save the output file with data textbox contents
                    Call rtbMessages.SaveFile(.FileName)
                    '// command completed, prompt user
                    MsgBox "YM Archive Saved Successfully", vbInformation + vbOKOnly, "Save Decoded Archive"
                End If
            End With
        End If
    Else    '// if data textbox is empty
        MsgBox "Please select an archive file to decode first", vbInformation + vbOKOnly, "Export To HTML"
    End If
'// error handler label
hell:
    If Err.Number = 0 Or Err.Number = 32755 Then
        Resume Next
    End If
End Sub
Private Sub cmdBrowse_Click()
    '// get the root directory by SHBrowseDialog api
    Dim strPath As String
    strPath = SHBrowseDialog("Selec source directory", Me)
    '// if selected directory is true
    If strPath <> "" Then
        '// if path string doesnt contain "\"
        If Right(strPath, 1) <> "\" Then
            strPath = strPath & "\"
        End If
        strFolder = strPath        '// update strFolder value
        txtPath.Text = strPath     '// update txtPath text
        '// load profiles again with the new provided path
        Call LoadProfiles(strFolder)
    End If
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    '// check for any previous instances of this application
    If PrevInstance = True Then
        MsgBox "Another instance of NexusYahoo!Decoder is running!", vbExclamation + vbOKCancel, "NexusYahoo!Decoder (Beta)"
        End '// end application
    End If
    
    '// assign the root path by reading it from app config file
    Call LoadSettings
    
    '// clear profiles treeview nodes
    tvwProfiles.Nodes.Clear
    rtbMessages.Text = ""   '// clear messages textbox
    
    '// populate the treeview with profiles
    Call LoadProfiles(strFolder)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '// if user clicked the X button on form
    If UnloadMode = 0 Then
        '// if permission to end application is false
        If MsgBox("Are you sure, you want to exit this program?", vbQuestion + vbOKCancel) = vbCancel Then
            Cancel = 1  '// cancel the unload process
        End If
    End If
End Sub

Private Sub Form_Resize()
    '// if we're not in minimized state
    If Me.WindowState <> vbMinimized Then
        '// check for the default size of main form
        If Me.Width < 8595 Or Me.Height < 6180 Then
            Me.Width = 8595
            Me.Height = 6180
        End If
        
        '// resize txtPath and cmdBrowse controls
        txtPath.Width = picPath.Width - lblPath.Width - cmdBrowse.Width - 200
        cmdBrowse.Left = txtPath.Left + txtPath.Width + 80
        
        '// resize profiles treeview
        With Me.tvwProfiles
            .Left = 80
            .Top = Me.cbMain.Height + 80
            .Height = (Me.ScaleHeight - Me.sbMain.Height) - 1150
        End With
        
        '// resize messages textbox
        With Me.rtbMessages
            .Left = (Me.tvwProfiles.Left + Me.tvwProfiles.Width) + 50
            .Top = Me.tvwProfiles.Top
            .Height = Me.tvwProfiles.Height
            .Width = Me.ScaleWidth - Me.tvwProfiles.Width - 250
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '// update screen mouse pointer
    Screen.MousePointer = vbHourglass
    
    '// unload and dispose every form in this application
    Dim objForm As Form
    For Each objForm In Forms
        Unload objForm
        Set objForm = Nothing
    Next objForm
    
    '// save app config (root directory path)
    Call SaveSettings
    '// update screen mouse pointer
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuBackup_Click()
    '// backup menu click
    frm_backup.Show vbModal, Me '// show backup form
End Sub

Private Sub mnuEnable_Click()
    '// profile enabler menu click
    frm_enabler.Show vbModal, Me    '// show profile enabler form
End Sub

Private Sub mnuFileExit_Click()
    '// exit menu click
    Unload Me   '// unload main form
End Sub

Private Sub mnuFileOpen_Click()
    '// file open menu click
    '// initialize CommonDialog props
    With Cdl
        .DialogTitle = "Open YM Archive File"
        .CancelError = True
        .Filter = "YM Archive Files|*.dat"
        .DefaultExt = "*.dat"
        .ShowOpen
        '// if input filename is specified
        If .FileName <> "" Then
            Dim strUsername As String   '// username string
            '// gather the username string from user's input
            strUsername = InputBox("Please enter your Yahoo! ID here:" & _
                vbCrLf & "(Entering invalid username will cause decoding to fail)", _
                "Enter Yahoo! ID", "inexuscore")
            '// if username string is provided
            If strUsername <> "" Then
                '// decode the input archive file by provided username string
                Call DecodeAsRTF(strUsername, .FileName)
            End If
        End If
    End With
End Sub

Private Sub mnuFileSave_Click()
    '// file save menu click
    '// if messages textbox is empty
    If Me.rtbMessages.Text = "" Then
        '// prompt user, then exit this procedure
        MsgBox "Please select an archive to decode first", vbInformation + vbOKOnly, "Save Decoded Archive"
        Exit Sub
    End If
    
    On Error GoTo hell
    
    '// initialize CommonDialog props
    With Cdl
        .DialogTitle = "Save Decoded Archive As..."
        .CancelError = True
        .Flags = &H2
        .Filter = "HTML Document(*.htm)|*.htm|Text File(*.txt)|*.txt|RTF Document(*.rtf)|*.rtf"
        .DefaultExt = "*.htm"
        .ShowSave
        '// if output filename is specified
        If .FileName <> "" Then
            '// get the selected node's data
            strNodeData = Split(tvwProfiles.SelectedItem.Key, ",")
            '// if selected node is a message node (archive files)
            If UBound(strNodeData) = 2 Then
                '// specifiy the selected file format,
                '// then save the selected archive file with that
                Select Case LCase(Right(.FileName, 3))
                    Case "htm"  '// HTML format
                        Call DecodeAsHTML(strNodeData(0), strFolder & strNodeData(0) & "\Archive\Messages\" & strNodeData(1) & "\" & strNodeData(2), .FileName)
                    Case "rtf"  '// RTF format
                        Call rtbMessages.SaveFile(.FileName)
                    Case "txt"  '// Text format
                        Dim X As Integer    '// free file handle
                        X = FreeFile    '// create a new free file
                        '// open output file for output access
                        Open .FileName For Output As #X
                            '// print the messages textbox as output file data
                            Print #X, rtbMessages.Text
                        Close #X    '// close output file
                End Select
            End If
            '// command completed, prompt user
            MsgBox "YM Archive Saved Successfully", vbInformation + vbOKOnly, "Save Decoded Archive"
        End If
    End With
'// error handler label
hell:
    If Err.Number = 0 Or Err.Number = 32765 Then
        Resume Next
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    '// about menu click
    frm_about.Show vbModal, Me  '// show about form
End Sub

Private Sub mnuHelpContents_Click()
    '// help menu click
    '// call ShellExecute and execute program's help file (html)
    ShellExecute Me.hwnd, "open", App.Path & "\Resources\help_contents.html", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub rtbMessages_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'// this function is not supported yet,
'// with the current decoding algorithm you have to
'// provide the username string for each archive file
'// in order to decode the data, hope i can fix this in near feature

'    Dim strUsername As String
'
'    If bLoading = False Then
'        If LCase(Right(Data.Files.Item(1), 3)) = "dat" Then
'            strUsername = InputBox("Enter the Yahoo! Username which this archive file belongs to:", "Drag & Drop Support", "inexuscore")
'
'            If Len(strUsername) > 0 Then
'                Call DecodeAsRTF(strUsername, Data.Files.Item(1))
'            End If
'        End If
'    End If
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    '// main toolbar button click
    '// check pressed button's key string
    Select Case Button.Key
        Case "Save" '// save button
            Call mnuFileSave_Click  '// call FileSave menu click
        Case "CopySelection"    '// copy selection button
            '// copy the selected text to clipboard
            '// if messages textbox is not empty
            If rtbMessages.Text <> "" Then
                '// if messages textbox has a selected string
                If rtbMessages.SelText <> "" Then
                    '// update screen mouse pointer
                    Screen.MousePointer = vbHourglass
                    '// copy selected string to clipboard
                    Clipboard.SetText rtbMessages.SelText
                    '// update screen mouse pointer
                    Screen.MousePointer = vbDefault
                Else    '// if theres no selection
                    MsgBox "No selections were found", vbInformation + vbOKOnly, "Copy Selection"
                End If
            Else    '// if messages textbox is empty
                MsgBox "Please select an archive to decode first", vbInformation + vbOKOnly, "Copy Selection"
            End If
        Case "Refresh"  '// refresh button
            '// get the selected node's data (profile, buddy, archive)
            strNodeData = Split(tvwProfiles.SelectedItem.Key, ",")
            '// if selected node is a message node (archive files)
            If UBound(strNodeData) = 2 Then
                '// decode the selected archive file as RTF format
                Call DecodeAsRTF(strNodeData(0), strFolder & strNodeData(0) & "\Archive\Messages\" & strNodeData(1) & "\" & strNodeData(2))
            End If
        Case "FindText" '// find text button
            '// not provided yet
        Case "Backup"   '// backup button
            Call mnuBackup_Click    '// call Backup menu click
        Case "ProfileEnabler"   '// profile enabler button
            Call mnuEnable_Click    '// call Enable menu click
        Case "About"    '// about button
            Call mnuHelpAbout_Click '//call HelpAbout menu click
        Case "Exit" '// exit button
            Call mnuFileExit_Click  '// call FileExit menu click
    End Select
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    '// check clicked button menu key string
    Select Case ButtonMenu.Key
        Case "ExportHTML"   '// export as html menu
            Call ExportToHTML
        Case "ExportText"   '// export as text menu
            Call ExportToText
        Case "ExportRTF"    '// export as rtf menu
            Call ExportToRTF
    End Select
End Sub

Private Sub tvwProfiles_NodeClick(ByVal Node As MSComctlLib.Node)
    '// profiles treeview node click
    '// if we're not decoding a message already
    If bLoading = False Then
        '// get the selected node's data (profile, buddy, archive)
        strNodeData = Split(Node.Key, ",")
        '// if selected node is a message node (archive files)
        If UBound(strNodeData) = 2 Then
            '// decode the selected archive file as rtf format
            Call DecodeAsRTF(strNodeData(0), strFolder & strNodeData(0) & "\Archive\Messages\" & strNodeData(1) & "\" & strNodeData(2))
        End If
    Else    '// if we're decoding an archive file already
        Call MsgBox("Please wait for current messages to finish loading.", vbInformation + vbOKOnly, "Nexus Y!Decoder")
    End If
End Sub

