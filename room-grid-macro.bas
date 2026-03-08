'================================================================
' ROOM MANAGEMENT GRID — VBA MACRO
'================================================================
' Reads SessionDB and populates "Room Management" sheet with
' session titles in the matching Day/Time/Room cells.
'
' To install: Alt+F11 > paste into existing Module1
'             (add below your existing RefreshAllDays code)
' To run:     Alt+F8 > RefreshRoomGrid > Run
'
' To add a new room: insert a column in the Room Management
' sheet, type the room name in row 2 (matching SessionDB
' column G exactly), type the role in row 3, then re-run
' this macro. It picks up new columns automatically.
'================================================================

Sub RefreshRoomGrid()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsGrid As Worksheet
    Dim wsDB As Worksheet

    On Error Resume Next
    Set wsGrid = ThisWorkbook.Sheets("Room Management")
    On Error GoTo 0
    If wsGrid Is Nothing Then
        MsgBox "Sheet 'Room Management' not found.", vbExclamation
        Exit Sub
    End If

    Set wsDB = ThisWorkbook.Sheets("SessionDB")

    Dim firstRoomCol As Long: firstRoomCol = 3
    Dim lastRoomCol As Long: lastRoomCol = firstRoomCol
    Do While wsGrid.Cells(2, lastRoomCol).Value <> ""
        lastRoomCol = lastRoomCol + 1
    Loop
    lastRoomCol = lastRoomCol - 1

    Dim roomCount As Long: roomCount = lastRoomCol - firstRoomCol + 1

    Dim roomNames() As String
    ReDim roomNames(firstRoomCol To lastRoomCol)
    Dim rc As Long
    For rc = firstRoomCol To lastRoomCol
        Dim rawName As String
        rawName = CStr(wsGrid.Cells(2, rc).Value & "")
        Dim brPos As Long
        brPos = InStr(rawName, vbLf)
        If brPos > 0 Then
            roomNames(rc) = Left(rawName, brPos - 1)
        Else
            roomNames(rc) = rawName
        End If
    Next rc

    Dim dataFirstRow As Long: dataFirstRow = 4
    Dim dataLastRow As Long: dataLastRow = 103

    Dim clearR As Long, clearC As Long
    For clearR = dataFirstRow To dataLastRow
        For clearC = firstRoomCol To lastRoomCol
            wsGrid.Cells(clearR, clearC).Value = ""
            wsGrid.Cells(clearR, clearC).Interior.ColorIndex = xlNone
            wsGrid.Cells(clearR, clearC).Font.Color = RGB(30, 41, 59)
            wsGrid.Cells(clearR, clearC).Font.Bold = False
        Next clearC
    Next clearR

    Dim dbLastRow As Long
    dbLastRow = wsDB.Cells(wsDB.Rows.Count, 2).End(xlUp).Row

    Dim sr As Long
    For sr = 2 To dbLastRow
        Dim sDay As Long: sDay = wsDB.Cells(sr, 2).Value
        If sDay < 1 Or sDay > 5 Then GoTo NextRoomSession

        Dim sStart As String: sStart = CStr(wsDB.Cells(sr, 3).Value & "")
        Dim sEnd As String: sEnd = CStr(wsDB.Cells(sr, 4).Value & "")
        Dim sTitle As String: sTitle = CStr(wsDB.Cells(sr, 5).Value & "")
        Dim sLead As String: sLead = CStr(wsDB.Cells(sr, 6).Value & "")
        Dim sRoom As String: sRoom = CStr(wsDB.Cells(sr, 7).Value & "")

        Dim targetCols() As Long
        Dim targetCount As Long: targetCount = 0
        ReDim targetCols(1 To roomCount)

        If sRoom = "All Rooms" Then
            For rc = firstRoomCol To lastRoomCol
                If roomNames(rc) <> "Main Hall" And roomNames(rc) <> "Atrium" And roomNames(rc) <> "Terrace" Then
                    If wsGrid.Cells(3, rc).Value = "Primary" Then
                        targetCount = targetCount + 1
                        targetCols(targetCount) = rc
                    End If
                End If
            Next rc
        Else
            For rc = firstRoomCol To lastRoomCol
                If roomNames(rc) = sRoom Then
                    targetCount = 1
                    targetCols(1) = rc
                    Exit For
                End If
            Next rc
        End If

        If targetCount = 0 Then GoTo NextRoomSession

        Dim sh As Long: sh = CLng(Left(sStart, 2))
        Dim sm As Long: sm = CLng(Mid(sStart, 4, 2))
        Dim eh As Long: eh = CLng(Left(sEnd, 2))
        Dim em As Long: em = CLng(Mid(sEnd, 4, 2))
        Dim tMin As Long: tMin = sh * 60 + sm
        Dim eMin As Long: eMin = eh * 60 + em

        Do While tMin < eMin
            Dim ts As String
            ts = Format(tMin \ 60, "00") & ":" & Format(tMin Mod 60, "00")

            Dim gridRow As Long: gridRow = 0
            Dim searchR As Long
            For searchR = dataFirstRow To dataLastRow
                If wsGrid.Cells(searchR, 1).Value = "D" & sDay And CStr(wsGrid.Cells(searchR, 2).Value & "") = ts Then
                    gridRow = searchR
                    Exit For
                End If
            Next searchR

            If gridRow > 0 Then
                Dim tc As Long
                For tc = 1 To targetCount
                    Dim tCol As Long: tCol = targetCols(tc)
                    If wsGrid.Cells(gridRow, tCol).Value = "" Then
                        If ts = sStart Then
                            If sLead <> "" Then
                                wsGrid.Cells(gridRow, tCol).Value = sTitle & vbLf & "Lead: " & sLead
                            Else
                                wsGrid.Cells(gridRow, tCol).Value = sTitle
                            End If
                            wsGrid.Cells(gridRow, tCol).Font.Bold = True
                        Else
                            wsGrid.Cells(gridRow, tCol).Value = ChrW(8595) & " " & sTitle
                        End If
                        wsGrid.Cells(gridRow, tCol).Interior.Color = RGB(219, 234, 254)
                        wsGrid.Cells(gridRow, tCol).Font.Color = RGB(30, 64, 175)
                        wsGrid.Cells(gridRow, tCol).Font.Size = 8
                        wsGrid.Cells(gridRow, tCol).WrapText = True
                    End If
                Next tc
            End If

            tMin = tMin + 30
        Loop

NextRoomSession:
    Next sr

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim bookedCount As Long: bookedCount = 0
    For clearR = dataFirstRow To dataLastRow
        For clearC = firstRoomCol To lastRoomCol
            If wsGrid.Cells(clearR, clearC).Value <> "" Then bookedCount = bookedCount + 1
        Next clearC
    Next clearR

    MsgBox "Room grid refreshed." & vbLf & roomCount & " rooms detected." & vbLf & bookedCount & " slots booked.", vbInformation, "Room Grid Refresh"
End Sub
