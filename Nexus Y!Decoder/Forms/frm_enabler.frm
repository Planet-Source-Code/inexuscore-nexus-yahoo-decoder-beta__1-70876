VERSION 5.00
Begin VB.Form frm_enabler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahoo! Profile Enabler"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_enabler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnable 
      Caption         =   "&Enable"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6060
      TabIndex        =   5
      Top             =   2760
      Width           =   1275
   End
   Begin NexysYDecoder.uc_ThreeDLine uc_ThreeDLine1 
      Height          =   45
      Left            =   120
      TabIndex        =   3
      Top             =   2580
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   79
      LineColour      =   0
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   3435
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frm_enabler.frx":7D42
      Height          =   1395
      Left            =   3720
      TabIndex        =   4
      Top             =   1020
      Width           =   3615
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label2"
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3660
      Picture         =   "frm_enabler.frx":7E44
      Top             =   420
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3660
      Picture         =   "frm_enabler.frx":870E
      Top             =   420
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Availabel profiles :"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1320
   End
End
Attribute VB_Name = "frm_enabler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mFS As FileSystemObject  '// filesystem object
Dim mFolder As Folder        '// folder object
Dim mSubFolder As Folder     '// sub folder object
Dim mReg As cls_Registry     '// registry class object
Private Sub cmdCancel_Click()
    Unload Me   '// unload profile enabler form
End Sub

Private Sub cmdEnable_Click()
    '// if selected profile is true
    If List1.Text <> "" Then
        Set mReg = New cls_Registry '// new instance of reg object
        '// initialize reg object props
        With mReg
            .ClassKey = HKEY_CURRENT_USER
            '// get the selected profile's reg key
            .SectionKey = "Software\yahoo\pager\profiles\" & _
                List1.Text & "\Archive"
            '// if selected profile reg key exists
            If .KeyExists = True Then
                '// set the AutoDelete value to 1 (true)
                .ValueKey = "AutoDelete"
                .ValueType = REG_DWORD
                .Value = 0
                '// set the Enabled value to 1 (true)
                .ValueKey = "Enabled"
                .ValueType = REG_DWORD
                .Value = 1
                '// set the Initialized value to 1 (true)
                .ValueKey = "Initialized"
                .ValueType = REG_DWORD
                .Value = 1
                
                '// update images visibility (lock, unlock images)
                Image1.Visible = False
                Image2.Visible = True
                '// update the status message label
                lblStatus.Caption = "Archiving is enabled for this profile"
                '// command completed, prompt uder
                MsgBox "Archiving option enabled for this profile", _
                    vbInformation + vbOKOnly, "Profile Enabler"
            Else
                '// if selected profile's reg key doesnt exist
                MsgBox "Selected profile has no correct registry keys!", _
                    vbCritical + vbOKCancel, "Profile Enabler"
            End If
        End With
    End If
End Sub

Private Sub Form_Load()
    List1.Clear '// clear profiles listbox
    '// update status message label
    lblStatus.Caption = "Idle"
    '// update images visivility (lock, unlock images)
    Image1.Visible = True
    Image2.Visible = False

    '// get the profile's root path from main form (txtPath.Text)
    Dim strPath As String
    strPath = frm_main.txtPath.Text
    '// get the root folder's handle
    Set mFS = New FileSystemObject
    Set mFolder = mFS.GetFolder(strPath)
    '// search for every profile folders and
    '// add the names into the listbox
    For Each mSubFolder In mFolder.SubFolders
        List1.AddItem mSubFolder.Name
    Next mSubFolder
    '// dispose filesystem objects
    Set mFS = Nothing
    Set mFolder = Nothing
End Sub

Private Sub List1_Click()
    '// if selected profile is true
    If List1.Text <> "" Then
        Set mReg = New cls_Registry '// new instance of reg object
        '// initialize reg object props
        With mReg
            .ClassKey = HKEY_CURRENT_USER
            '// get the selected profile's reg key
            .SectionKey = "Software\yahoo\pager\profiles\" & _
                List1.Text & "\Archive"
            '// if reg key exists
            If .KeyExists = True Then
                '// update images visibility (lock, unlock images)
                Image1.Visible = False
                Image2.Visible = True
                '// update the status mesasge label
                lblStatus.Caption = "Archiving is enabled for this profile"
            Else
                '// update images visibility (lock, unlock images)
                Image1.Visible = True
                Image2.Visible = False
                '// update the status message label
                lblStatus.Caption = "Archiving is not enabled for this profile"
            End If
        End With
    End If
End Sub
