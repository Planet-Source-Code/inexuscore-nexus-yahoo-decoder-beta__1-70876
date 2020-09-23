VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Archive Data"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_backup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6420
      TabIndex        =   7
      Top             =   4500
      Width           =   1395
   End
   Begin NexysYDecoder.uc_ThreeDLine uc_ThreeDLine1 
      Height          =   45
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   79
      LineColour      =   0
   End
   Begin VB.CheckBox chkStorePath 
      Caption         =   "Store full path for files and folders"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   3900
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CheckBox chkBackupAll 
      Caption         =   "Backup all archives for selected profile"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3900
      Width           =   3435
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   2355
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4154
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Make The Backup"
      Height          =   375
      Left            =   4740
      TabIndex        =   6
      Top             =   4500
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   4980
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   3420
      Width           =   855
   End
   Begin VB.TextBox txtOutput 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3420
      Width           =   6735
   End
   Begin VB.ComboBox cboProfiles 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Output File :"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Profile :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frm_backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strRootPath As String           '// contains profiles root path
Dim mFS As New FileSystemObject     '// file system object
Dim mFolder As Folder               '// folder object
Dim mSubFolder As Folder            '// sub folder object
Dim mFile As File                   '// file object
Private Sub cboProfiles_Click()
    Dim strProfile As String    '// selected profile name
    Dim strPath As String       '// profile's archive path
    Dim itmX As ListItem        '// listitem object
    
    strProfile = cboProfiles.Text   '// get the selected profile
    '// get the selected profile's archive path
    strPath = strRootPath & strProfile & "\Archive\Messages\"
    
    '// if aarchive path exists
    If Dir(strPath, vbDirectory) <> "" Then
        Dim mSize As Long   '// total size of archive files
        
        '// get the archive folder's handle
        Set mFS = New FileSystemObject  '// create a new instance
        Set mFolder = mFS.GetFolder(strPath)
        '// search for all buddy folders, then add the names
        '// and total size of files in listview
        For Each mSubFolder In mFolder.SubFolders
            '// add the buddy folder name
            Set itmX = lvwData.ListItems.Add(, , mSubFolder.Name)
            '// add the buddy archive files count
            itmX.SubItems(1) = mSubFolder.Files.Count
            mSize = 0   '// reset total size
            '// calculate total size of archive files for
            '// earch buddy folder
            For Each mFile In mSubFolder.Files
                mSize = mSize + mFile.Size  '// increase the total size
            Next mFile
            '// format total size into KB (##.##)
            itmX.SubItems(2) = Format(mSize / 1024, "00.00")
        Next mSubFolder
        '// dispose FileSystem objects
        Set mFS = Nothing
        Set mFolder = Nothing
        Set mSubFolder = Nothing
        Set mFile = Nothing
    End If
End Sub

Private Sub cmdBackup_Click()
    '// if output filename is unspecified
    If txtOutput.Text = "" Then
        MsgBox "Please specify the output filename first.", _
            vbInformation + vbOKCancel, "Output Filename"
        Exit Sub    '// exit this procedure
    End If
    
    '// if there's a selected profile and listview is not empty
    If cboProfiles.Text <> "" And lvwData.ListItems.Count <> 0 Then
        Dim strPath As String   '// archive path
        Dim mZip As New Zip     '// XZip object
        
        '// if Backup All Archives is selected
        If chkBackupAll.Value = vbChecked Then
            '// get the archives path
            strPath = strRootPath & cboProfiles.Text & "\Archive\Messages"
            
            '// update screen mouse pointer
            Screen.MousePointer = vbHourglass
            
            '// get the archive folder's handle
            Set mFS = New FileSystemObject
            Set mFolder = mFS.GetFolder(strPath)
            '// search for every buddy archive files
            For Each mSubFolder In mFolder.SubFolders
                For Each mFile In mSubFolder.Files
                    '// add each archive file into the zip archive
                    mZip.Pack mFile.Path, txtOutput.Text, chkStorePath.Value
                Next mFile
            Next mSubFolder
            
            '// update screen mouse pointer
            Screen.MousePointer = vbDefault
        Else
            '// get the selected buddy archive path
            strPath = strRootPath & cboProfiles.Text & _
                "\Archive\Messages\" & lvwData.SelectedItem.Text
            
            '// update screen mouse pointer
            Screen.MousePointer = vbHourglass
            
            '// get the archive folder's handle
            Set mFS = New FileSystemObject
            Set mFolder = mFS.GetFolder(strPath)
            '// search for every archive files
            For Each mFile In mFolder.Files
                '// add each archive file to the zip archive
                mZip.Pack mFile.Path, txtOutput.Text, chkStorePath.Value
            Next mFile
            '// update screen mouse pointer
            Screen.MousePointer = vbDefault
        End If
        
        '// dispose FileSystem objects and the XZip object
        Set mFS = Nothing
        Set mFolder = Nothing
        Set mSubFolder = Nothing
        Set mFile = Nothing
        Set mZip = Nothing
        
        '// command completed, prompt user
        MsgBox "Backup archive created successfully at: " & vbCrLf & _
            txtOutput.Text, vbInformation + vbOKOnly, "Backup Archives"
    End If
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo hell
    
    '// initialize CommonDialog control props
    With Cdl
        .CancelError = True
        .DialogTitle = "Save backup file as..."
        .Filter = "Zip Archive Files(*.zip)|*.zip"
        .Flags = &H2
        .ShowSave
        
        '// if output filename is specified
        If .FileName <> "" Then
            '// update txtOutput text with specified filename
            txtOutput.Text = .FileName
        End If
    End With
'// error handler label
hell:
    If Err.Number = 0 Or Err.Number = 32755 Then
        Resume Next
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me   '// unload backup form
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    '// initialize lvwData listview
    With lvwData
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .FlatScrollBar = False
        
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        .ColumnHeaders.Add , , "BuddyName", 5200
        .ColumnHeaders.Add , , "Archive Files", 1200
        .ColumnHeaders.Add , , "Size KB", .Width - 6800
    End With
    
    '// get the profiles root path from main form (txtPath.Text)
    strRootPath = frm_main.txtPath.Text
    '// get the root folder's handle
    Set mFS = New FileSystemObject
    Set mFolder = mFS.GetFolder(strRootPath)
    '// clear profiles combobox
    cboProfiles.Clear
    '// search for all existed profile folders,
    '// and add the names into the combobox
    For Each mSubFolder In mFolder.SubFolders
        cboProfiles.AddItem mSubFolder.Name
    Next mSubFolder
    '// dispose FileSystem objects
    Set mFS = Nothing
    Set mFolder = Nothing
End Sub

Private Sub txtOutput_KeyPress(KeyAscii As Integer)
    Call cmdBrowse_Click    '// show the SaveDialog for output filename
End Sub
