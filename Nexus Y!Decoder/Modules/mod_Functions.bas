Attribute VB_Name = "mod_Functions"
Option Explicit

Public Function PrevInstance() As Boolean
    '// Check for a window containing the InstanceCode
    '// If it is found, then return true (another instance is running)
    If FindWindow(vbNullString, ByVal InstanceCode) Then
        PrevInstance = True
        Exit Function
    End If
    
    '// Else, create the window with the InstanceCode
    CreateWindowEx 0&, "STATIC", InstanceCode, 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, 0&
    PrevInstance = False
End Function
Public Function SHBrowseDialog(ByVal Title As String, ByVal Form As Form) As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    szTitle = Title

    With tBrowseInfo
        .hwndOwner = Form.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        SHBrowseDialog = sBuffer
    End If
End Function

