Sub CheckTableSize()
    ' Table Size Analysis
    Dim DB As DAO.Database, NewDB As String, T As DAO.TableDef, SizeAft As Long, _
        SizeBef As Long, RST As DAO.Recordset, F As Boolean, RecCnt As Long
    
    Const StTable As String = "_Tables"
    
    Set DB = CurrentDb
    
    NewDB = Left(DB.Name, InStrRev(DB.Name, "\")) & Replace(Str(Now), ":", "-") & " " & _
        Mid(DB.Name, InStrRev(DB.Name, "\") + 1, Len(DB.Name))
    Application.DBEngine.CreateDatabase NewDB, DB_LANG_GENERAL
    
    F = False
    For Each T In DB.TableDefs
        If T.Name = StTable Then
            F = True: Exit For
        End If
    Next T
    If F Then
        DB.Execute "DELETE FROM " & StTable, dbFailOnError
    Else
        DB.Execute "CREATE TABLE " & StTable & _
            " (tblName TEXT(255), tblRecords LONG, tblSize LONG);", dbFailOnError
    End If
    
    For Each T In DB.TableDefs
        ' Exclude system tables:
        If Not T.Name Like "MSys*" And T.Name <> StTable And T.Name <> "motiv" Then
            RecCnt = T.RecordCount
            ' If it's linked table:
            If RecCnt = -1 Then RecCnt = DCount("*", T.Name)
            If RecCnt > 0 Then DB.Execute "INSERT INTO " & StTable & _
                " (tblName, tblRecords) " & _
                "VALUES ('" & T.Name & "', " & RecCnt & ")", dbFailOnError
        End If
    Next T
    
    Set RST = DB.OpenRecordset("SELECT * FROM " & StTable, dbOpenDynaset)
    If RST.RecordCount > 0 Then
        Do Until RST.EOF
            Debug.Print "Processing table " & RST("tblName") & "..."
            SizeBef = FileLen(NewDB)
            DB.Execute ("SELECT * " & _
            "INTO " & RST("tblName") & " IN '" & NewDB & "' " & _
            "FROM " & RST("tblName")), dbFailOnError
            SizeAft = FileLen(NewDB) - SizeBef
            RST.Edit
                RST("tblSize") = SizeAft
            RST.Update
            Debug.Print "    size = " & SizeAft
            RST.MoveNext
        Loop
    Else
        Debug.Print "No tables found!"
    End If
    RST.Close: Set RST = Nothing
    
    Debug.Print ">>> Done! <<<"
    MsgBox "Done!", vbInformation + vbSystemModal, "CheckTableSize"
    Kill NewDB
    Set DB = Nothing
    
    DoCmd.OpenTable StTable, acViewNormal, acReadOnly
End Sub
