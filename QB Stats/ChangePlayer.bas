Attribute VB_Name = "ChangePlayer"
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim name_playerr As String
    Dim first_name As String
    Dim last_name As String
    Dim ws_name As String
    Dim import_row As Integer

    If Target.Address = "$B$2" Then
        ws_name = ActiveSheet.Name
        With Sheets(ws_name)
            Application.EnableEvents = False
            name_player = Sheets(ws_name).Cells(2, 2).Value
            first_name = LCase(Trim(Mid(name_player, 1, InStr(1, name_player, " ") - 1)))
            last_name = LCase(Trim(Mid(name_player, InStr(1, name_player, " ") + 1)))
            name_table = StrConv(first_name, vbProperCase) & "_" & StrConv(last_name, vbProperCase)
           
            delete_tables ws_name, name_table


            import_row = import_row_find(ws_name)
            player_stats_basic import_row, name_table

            import_row = import_row_find(ws_name)
            player_stats_per_game import_row, name_table
           
            import_row = import_row_find(ws_name)
            player_stats_16_game import_row, name_table

            Application.EnableEvents = True
        End With
    End If
End Sub

Private Sub delete_tables(ByVal ws_name, ByVal name_table)
    Dim last_used_row As Integer
    Dim last_used_col As Integer
    Dim wkb_name As String
    Dim wb_connection As Object

    wkb_name = ActiveWorkbook.Name

    With Sheets(ws_name)
        last_used_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        last_used_col = .Cells(1, .Columns.Count).End(xlToLeft).Column
        .Range(.Cells(4, 1), .Cells(last_used_row, last_used_col)).EntireRow.Delete
    End With

    'For Each wb_connection In Workbooks(wkb_name).Connections
    '    If InStr(1, wb_connection.Name, name_table) Then
    '        ActiveWorkbook.Connections(wb_connection.Name).Delete
    '    End If
    'Next


End Sub

Private Function import_row_find(ByVal ws_name)
    Dim last_used_row As Integer

    With Sheets(ws_name)
        last_used_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        import_row_find = last_used_row + 2
    End With
End Function

Private Sub player_stats_basic(ByVal import_row, ByVal name_table)
    Application.CutCopyMode = False
    Workbooks("QBs.xlsm").Connections.Add2 _
        "WorksheetConnection_QBs.xlsm!" & name_table & "", "", _
        "WORKSHEET;C:\Python36\projects\Fantasy Football Projections\Basic Stats\QBs.xlsm" _
        , "QBs.xlsm!" & name_table & "", 7, True, False
    With ActiveSheet.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook. _
        Connections("WorksheetConnection_QBs.xlsm!" & name_table & ""), _
        Destination:=Range("$A$" & import_row)).TableObject
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = 1
        .AdjustColumnWidth = True
        .ListObject.DisplayName = "Player_Base_Stats"
        .Refresh
    End With
    ActiveWindow.Zoom = 56
End Sub
Private Sub player_stats_per_game(ByVal import_row, ByVal name_table)
    Application.CutCopyMode = False
    Workbooks("QBs.xlsm").Connections.Add2 _
        "WorksheetConnection_QBs.xlsm!" & name_table & "_per_game", "", _
        "WORKSHEET;C:\Python36\projects\Fantasy Football Projections\Basic Stats\QBs.xlsm" _
        , "QBs.xlsm!" & name_table & "_per_game", 7, True, False
    With ActiveSheet.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook. _
        Connections("WorksheetConnection_QBs.xlsm!" & name_table & "_per_game"), _
        Destination:=Range("$A$" & import_row)).TableObject
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = 1
        .AdjustColumnWidth = True
        .ListObject.DisplayName = "Player_Per_Game_Stats"
        .Refresh
    End With
End Sub
Private Sub player_stats_16_game(ByVal import_row, ByVal name_table)
    Application.CutCopyMode = False
    Workbooks("QBs.xlsm").Connections.Add2 _
        "WorksheetConnection_QBs.xlsm!" & name_table & "_16_game", "", _
        "WORKSHEET;C:\Python36\projects\Fantasy Football Projections\Basic Stats\QBs.xlsm" _
        , "QBs.xlsm!" & name_table & "_16_game", 7, True, False
    With ActiveSheet.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook. _
        Connections("WorksheetConnection_QBs.xlsm!" & name_table & "_16_game"), _
        Destination:=Range("$A$" & import_row)).TableObject
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = 1
        .AdjustColumnWidth = True
        .ListObject.DisplayName = "Player_16_Game_Stats"
        .Refresh
    End With
End Sub



