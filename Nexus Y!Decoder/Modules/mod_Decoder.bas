Attribute VB_Name = "mod_Decoder"
Option Explicit

Private Function DecodeTime(ByVal inputTime As String) As String
    '// this function decodes the DateTime string
    Dim lngDate As Date
    Dim lngDateBase As Date
    Dim lngDateDiff As Long
    Dim lngDateTimeZone As Long
    Dim lngSeconds As Double
    
    lngDateBase = DateSerial(1970, 1, 1)
    lngDateDiff = 631152000
    lngDateTimeZone = 13
    
    lngSeconds = Asc(Mid$(inputTime, 1, 1))
    lngSeconds = lngSeconds + (Asc(Mid$(inputTime, 2, 1)) * 256#)
    lngSeconds = lngSeconds + (Asc(Mid$(inputTime, 3, 1)) * (256# ^ 2))
    lngSeconds = lngSeconds + (Asc(Mid$(inputTime, 4, 1)) * (256# ^ 3))
    lngSeconds = lngSeconds - lngDateDiff
    lngDate = DateAdd("s", lngSeconds, lngDateBase)
    lngDate = DateAdd("s", lngDateDiff, lngDate)
    lngDate = DateAdd("h", lngDateTimeZone, lngDate)
    
    DecodeTime = Mid(lngDate, 11, Len(lngDate))
End Function
Public Sub DecodeAsRTF(ByVal strUsername As String, ByVal strFile As String)
    Dim intFreeFile As Integer
    Dim strBuffer As String
    Dim strArray() As String
    Dim lngCount As Long
    Dim intLen As Integer
    Dim boolSent As Boolean
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim strSender As String
    Dim strTime As String
    Dim strDateTime As String
    
    On Error GoTo hell
    
    '// update screen mouse pointer
    Screen.MousePointer = vbHourglass
    '// update main form menus
    frm_main.mnuFileSave.Enabled = False
    '// update main form's toolbar buttons
    With frm_main.tbrMain
        .Buttons("Save").Enabled = False
        .Buttons("Export").Enabled = False
        .Buttons("CopySelection").Enabled = False
        .Buttons("Refresh").Enabled = False
        .Buttons("FindText").Enabled = False
    End With
    '// if selected archive file exists
    If Dir(strFile, vbNormal) <> "" Then
        '// set loading variable to true.
        '// this is used for when users attempt
        '// to open another archive while one is already loading, it won't let them.
        bLoading = True
        '// load file contents into strBuffer for processing
        intFreeFile = FreeFile
        '// open the archive file for binary access
        Open strFile For Binary As intFreeFile
            strBuffer = Space(LOF(intFreeFile))
            Get #intFreeFile, 1, strBuffer
        Close #intFreeFile '// close the archive file
        '// split file contents into array to process. Best method to me
        '// at the time was by splitting with delimeter of 3 null bytes.
        strArray = Split(strBuffer, String(3, vbNullChar))

        '// clear main form's messages textbox
        frm_main.rtbMessages.Text = vbNullString
        '// loop through every index in our array
        Do Until lngCount = UBound(strArray) + 1
            '// if the current item has an item after it and the current item
            '// contains data then...
            If lngCount + 1 <= UBound(strArray) And Len(strArray(lngCount)) > 0 Then
                '//gather the first 3 bits of XOR'd message, in order to
                '// decode the date-time
                strTime = strArray(0) + strArray(1) + strArray(2)
                strDateTime = DecodeTime(strTime)
                '// after testing it seems i can decide sent from received messages
                '// by the length of the item before the actual XOR'd message so we'll
                '// set our boolean variable appropriately
                If Len(strArray(lngCount)) = 1 Then
                    boolSent = False
                ElseIf Len(strArray(lngCount)) = 2 Then
                    boolSent = True
                End If
                '// after testing it seems the length of the XOR'd message is within
                '// the previous item (item that tells us sent/recv) as the ASCII
                '// value so we'll strip the null bytes and set out intLen variables
                '// appropriately.
                intLen = Asc(Replace(strArray(lngCount), vbNullChar, ""))
                '// if the length of the string is greater than 1 then...
                If intLen > 1 Then
                    '// if out intLen is the same as the length of the next item (which
                    '// is our XOR'd message) then we "have a winner!" >:)
                    If intLen = Len(strArray(lngCount + 1)) Then
                        '// set out intCount to 1 (the start of our username) so we can
                        '// begin XOR'ing the message.
                        intCount = 1
                        '// start our loop from the start and end of our XOR'd message in
                        '// our array
                        For intLoop = 1 To Len(strArray(lngCount + 1))
                            '// replace the current char with the XOR'd char
                            Mid(strArray(lngCount + 1), intLoop, 1) = Chr(Asc(Mid(strArray(lngCount + 1), intLoop, 1)) Xor Asc(Mid(strUsername, intCount, 1)))
                            '// increase our count so it'll use the next char in the username
                            '// to XOR the char in the XOR'd message
                            intCount = intCount + 1
                            '// if the count is beyond the length of our username, set the
                            '// count at 1 so it will reuse the username to XOR
                            If intCount > Len(strUsername) Then
                                intCount = 1
                            End If
                        Next intLoop
                        '//  our XOR'd message is now XOR'd back to the original message, so we'll
                        '//  add it to the RichTextBox
                        If boolSent = True Then
                            strSender = strUsername
                        Else
                            strSender = strNodeData(1)
                        End If
                        '// remove some extra tags created by yahoo! text editor
                        '// some tags like <ding>, text styling and more
                        Dim intMe As Integer
                        strOutput = strArray(lngCount + 1)

                        intMe = InStr(1, strOutput, "<ding>", vbTextCompare)

                        If intMe = 0 Then
                            intMe = InStr(1, strOutput, "</FADE>", vbTextCompare)

                            If intMe <> 0 Then
                                strOutput = Mid(strOutput, 1, Len(strOutput) - 7)
                            End If

                            intMe = InStr(1, strOutput, """>", vbTextCompare)

                            If intMe <> 0 Then
                                strOutput = Mid(strOutput, intMe + 2, Len(strOutput))
                            End If

                            intMe = InStr(1, strOutput, ">", vbTextCompare)

                            If intMe <> 0 Then
                                strOutput = Mid(strOutput, intMe + 1, Len(strOutput))
                            End If
                        End If
                        '// update main form's messages textbox
                        With frm_main.rtbMessages
                            .SelStart = Len(.Text)
                            .SelColor = IIf(boolSent = True, lngColorSent, lngColorRecv)
                            .SelBold = True
                            .SelText = "(" & strDateTime & ") " & strSender & ": "
                            .SelStart = Len(.Text)
                            .SelBold = False
                            .SelColor = vbBlack
                            .SelText = strOutput & vbCrLf
                            .SelStart = Len(.Text)
                        End With
                    End If
                End If
            End If
            '// update main form's status bar text
            frm_main.sbMain.Panels(1).Text = "Decoding messages {" & Int(((lngCount * 100) / UBound(strArray))) & "%...}"
            '// increment our array index count
            lngCount = lngCount + 1
            '// free up processing of other window-messages
            DoEvents
        Loop
        
        '// clear some variables at the end
        strBuffer = ""
        ReDim strArray(0)
        intLen = 0
        strSender = ""
        strTime = ""
        strDateTime = ""
        strOutput = ""
        '// set our boolean to false so the user can now load other archives
        bLoading = False
        '// update main form's status bar text
        frm_main.sbMain.Panels(1).Text = "Ready"
    End If
    
    '// update main menus, main toolbar buttons
    frm_main.mnuFileSave.Enabled = True
    '// update main form's toolbar buttons
    With frm_main.tbrMain
        .Buttons("Save").Enabled = True
        .Buttons("Export").Enabled = True
        .Buttons("CopySelection").Enabled = True
        .Buttons("Refresh").Enabled = True
        .Buttons("FindText").Enabled = True
    End With
    '// update screen mouse pointer
    Screen.MousePointer = vbDefault
'// error handler label
hell:
    If Err.Number = 6 Then  '// Overflow exception
        '// update screen mouse pointer
        Screen.MousePointer = vbDefault
        '// prompt user
        MsgBox "Invalid file or bad data. please try another archive files to decode", _
                vbExclamation + vbOKCancel, "Decoding Failed"
        '// resume to next command
        Resume Next
    End If
End Sub

Public Sub DecodeAsHTML(ByVal strUsername As String, ByVal strFile As String, ByVal strOutput As String)
    Dim intFreeFile As Integer
    Dim intOutput As Integer
    Dim strBuffer As String
    Dim strArray() As String
    Dim strArchiveName As String
    Dim lngCount As Long
    Dim intLen As Integer
    Dim boolSent As Boolean
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim strSender As String
    Dim strTime As String
    Dim strDateTime As String
    
    On Error GoTo hell
    
    '// update screen mouse pointer
    Screen.MousePointer = vbHourglass
    '// update mainf form menus
    frm_main.mnuFileSave.Enabled = False
    '// update main form's toolbar buttons
    With frm_main.tbrMain
        .Buttons("Save").Enabled = False
        .Buttons("Export").Enabled = False
        .Buttons("CopySelection").Enabled = False
        .Buttons("Refresh").Enabled = False
        .Buttons("FindText").Enabled = False
    End With
    
    '// if the archive file exists
    If Dir(strFile, vbNormal) <> "" Then
        '// Set loading variable to true. This is used for when users attempt
        '// to open another archive while one is already loading, it won't let them.
        bLoading = True
        '// Load file contents into strBuffer for processing
        intFreeFile = FreeFile
        intOutput = FreeFile
        '// open the archive file for binary access
        Open strFile For Binary As intFreeFile
            strBuffer = Space(LOF(intFreeFile))
            Get #intFreeFile, 1, strBuffer
        Close intFreeFile   '// close the archive file
        
        '// Open output html file and write the required tags and data here
        Open strOutput For Output As intOutput
            '// Get the archive file name by FileSystemObject
            Dim FS As New FileSystemObject
            Dim objFile As File
            
            Set objFile = FS.GetFile(strFile)
            strArchiveName = objFile.Name
            
            '// Dispose FileSystemObjects
            Set objFile = Nothing
            Set FS = Nothing
            
            '// Start writing the html tags and data here
            Print #intOutput, "<HTML>"
            Print #intOutput, "<HEAD>"
            Print #intOutput, "<STYLE>"
            Print #intOutput, "body{font-family:Trebuchet MS,sans-serif;padding-left:10px;padding-right:10px;font-size:10pt;margin:0 auto;}" & Chr(13) & _
                ".title{text-align:center;color:#090;font-size:18pt;border:0px;border-top:4px #000 solid;padding-top:5px;}" & Chr(13) & _
                ".author{text-align:center;border:0px;border-bottom:1px #999 solid;padding-bottom:5px;}" & Chr(13) & _
                ".convo-started {padding-top:10px;padding-bottom:10px;color:#090;}" & Chr(13) & _
                ".date{font-size:8pt; color:#300;}" & Chr(13) & _
                ".local{font-weight:bold; font-size:8pt;color:#888;font-family:Tahoma;}" & Chr(13) & _
                ".remote{font-weight:bold; font-size:8pt;color:#339;font-family:Tahoma;}" & Chr(13) & _
                ".msg{font-size:10pt; font-family:Arial;}" & Chr(13) & _
                ".buzz{color:#800;font-weight:bold;}"
            Print #intOutput, "</STYLE>"
            Print #intOutput, "<TITLE>"
            Print #intOutput, "Nexus Yahoo! Decoder - Archive: {" & strArchiveName & "}"
            Print #intOutput, "</TITLE>"
            Print #intOutput, "</HEAD>"
            Print #intOutput, "<BODY>"
            Print #intOutput, "<div class=""title"">Nexus Yahoo! Decoder (Beta)</div>"
            Print #intOutput, "<div class=""author"">INexusCore, Inc.<br/><a href=""mailto:inexuscore@gmail.com"">inexuscore@gmail.com</a><br/></div>"
        
        '// Split file contents into array to process. Best method to me
        '// at the time was by splitting with delimeter of 3 null bytes.
        strArray = Split(strBuffer, String(3, vbNullChar))
        
        '// Loop through every index in our array
        Do Until lngCount = UBound(strArray) + 1
            '// If the current item has an item after it and the current item
            '// Contains data then...
            If lngCount + 1 <= UBound(strArray) And Len(strArray(lngCount)) > 0 Then
                '// Gather the first 3 bits of XOR'd message, in order to
                '// decode the date-time
                strTime = strArray(0) + strArray(1) + strArray(2)
                strDateTime = DecodeTime(strTime)
                '// After testing it seems i can decide sent from received messages
                '// By the length of the item before the actual XOR'd message so we'll
                '// Set our boolean variable appropriately
                If Len(strArray(lngCount)) = 1 Then
                    boolSent = False
                ElseIf Len(strArray(lngCount)) = 2 Then
                    boolSent = True
                End If
                '// After testing it seems the length of the XOR'd message is within
                '// the previous item (item that tells us sent/recv) as the ASCII
                '// value so we'll strip the null bytes and set out intLen variables
                '// appropriately.
                intLen = Asc(Replace(strArray(lngCount), vbNullChar, ""))
                '// If the length of the string is greater than 1 then...
                If intLen > 1 Then
                    '// If out intLen is the same as the length of the next item (which
                    '// is our XOR'd message) then we "have a winner!" >:)
                    If intLen = Len(strArray(lngCount + 1)) Then
                        '// Set out intCount to 1 (the start of our username) so we can
                        '// begin XOR'ing the message.
                        intCount = 1
                        '// Start our loop from the start and end of our XOR'd message in
                        '// our array
                        For intLoop = 1 To Len(strArray(lngCount + 1))
                            '// Replace the current char with the XOR'd char
                            Mid(strArray(lngCount + 1), intLoop, 1) = Chr(Asc(Mid(strArray(lngCount + 1), intLoop, 1)) Xor Asc(Mid(strUsername, intCount, 1)))
                            '// Increase our count so it'll use the next char in the username
                            '// to XOR the char in the XOR'd message
                            intCount = intCount + 1
                            '// If the count is beyond the length of our username, set the
                            '// count at 1 so it will reuse the username to XOR
                            If intCount > Len(strUsername) Then
                                intCount = 1
                            End If
                        Next intLoop
                        '// Our XOR'd message is now XOR'd back to the original message, so we'll
                        '// add it to the RichTextBox
                        If boolSent = True Then
                            strSender = strUsername
                        Else
                            strSender = strNodeData(1)
                        End If
                                        
                        '// remove some extra tags created by yahoo! text editor
                        '// some tags like <ding>, text styling and more
                        Dim intMe As Integer
                        strOutput = strArray(lngCount + 1)

                        intMe = InStr(1, strOutput, "<ding>", vbTextCompare)

                        If intMe = 0 Then
                            intMe = InStr(1, strOutput, "</FADE>", vbTextCompare)

                            If intMe <> 0 Then
                                strOutput = Mid(strOutput, 1, Len(strOutput) - 7)
                            End If

                            intMe = InStr(1, strOutput, """>", vbTextCompare)

                            If intMe <> 0 Then
                                strOutput = Mid(strOutput, intMe + 2, Len(strOutput))
                            End If

                            intMe = InStr(1, strOutput, ">", vbTextCompare)

                            If intMe <> 0 Then
                                strOutput = Mid(strOutput, intMe + 1, Len(strOutput))
                            End If
                        End If

                        If boolSent = True Then
                            Print #intOutput, "<div><span class=""date"">(" & strDateTime & ")</span> <span class=""local"">" & strSender & ": " & "</span><span class=""msg"">" & strOutput & "</span></div>"
                        Else
                            Print #intOutput, "<div><span class=""date"">(" & strDateTime & ")</span> <span class=""remote"">" & strSender & ": " & "</span><span class=""msg"">" & strOutput & "</span></div>"
                        End If
                    End If
                End If
            End If
            '// update main form's status bar text
            frm_main.sbMain.Panels(1).Text = "Decoding messages {" & Int(((lngCount * 100) / UBound(strArray))) & "%...}"
            '// Increment our array index count
            lngCount = lngCount + 1
            '// Free up processing of other window-messages
            DoEvents
        Loop
        
        '// Write the ending html tags and close the html file here
        Print #intOutput, "<br/><br/>"
        Print #intOutput, "</BODY>"
        Print #intOutput, "</HTML>"
        Close intOutput
        '// Clear some variables at the end
        strBuffer = ""
        ReDim strArray(0)
        intLen = 0
        strSender = ""
        strTime = ""
        strDateTime = ""
        strOutput = ""
        '// Set our boolean to false so the user can now load other archives
        bLoading = False
    End If
    
    '// Update main menus, main toolbar buttons
    frm_main.mnuFileSave.Enabled = True
    '// update main form's toolbar buttons
    With frm_main.tbrMain
        .Buttons("Save").Enabled = True
        .Buttons("Export").Enabled = True
        .Buttons("CopySelection").Enabled = True
        .Buttons("Refresh").Enabled = True
        .Buttons("FindText").Enabled = True
    End With
    '// update screen mouse pointer
    Screen.MousePointer = vbDefault
'// error handler label
hell:
    If Err.Number = 6 Then  '// Overflow exception
        '// update screen mouse pointer
        Screen.MousePointer = vbDefault
        '// prompt user
        MsgBox "Invalid file or bad data. please try another archive files to decode", _
                vbExclamation + vbOKCancel, "Decoding Failed"
        '// resume to next command
        Resume Next
    End If
End Sub

