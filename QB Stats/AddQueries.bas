Attribute VB_Name = "AddQueries"
Option Explicit

Public Sub add_nfl_com_queries()
    Dim first_name As String
    Dim last_name As String
    Dim name_player As String
    Dim name_table As String
    Dim source_str As String
    Dim wp_str As String
    Dim main_sheet As String
    
    
    Dim start_row As Integer
    Dim last_row As Integer
    Dim row_count As Integer
    Dim last_row_qs As Integer
    Dim import_row As Integer
    Dim last_col_qs As Integer
    
    main_sheet = ActiveSheet.Name
    start_row = Selection.Rows(1).Row
    last_row = start_row + Selection.Rows.Count - 1
    Application.ScreenUpdating = False
    For row_count = start_row To last_row
        name_player = Sheets(main_sheet).Cells(row_count, 1).Value
        first_name = LCase(Trim(Mid(name_player, 1, InStr(1, name_player, " ") - 1)))
        last_name = LCase(Trim(Mid(name_player, InStr(1, name_player, " ") + 1)))
        wp_str = Sheets(main_sheet).Cells(row_count, 2).Value
        source_str = "Source = Web.Page(Web.Contents(" & Chr(34) & wp_str & Chr(34) & ")),"
        name_table = StrConv(first_name, vbProperCase) & "_" & StrConv(last_name, vbProperCase)
        add_player_query source_str, name_table
        
        ActiveSheet.Name = name_table
        
        add_per_game_table name_table

        With ActiveSheet
            .Columns("A:Z").AutoFit
        End With
        
    Next
    Application.ScreenUpdating = True
End Sub
Private Sub add_player_query(ByVal source_str, ByVal name_table)
    ActiveWorkbook.Queries.Add Name:="Table 1", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & Chr(9) & source_str & Chr(13) & "" & Chr(10) & "    Data1 = Source{1}[Data]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Data1,{{""Header"", type text}, {""Season"", type text}, {""Team"", type text}, {""G"", type text}, {""GS"", type text}, {""Passing Comp"", type number}, {""Passing Att"", type" & _
        " number}, {""Passing Pct"", type number}, {""Passing Yds"", type number}, {""Passing Avg"", type number}, {""Passing TD"", Int64.Type}, {""Passing Int"", Int64.Type}, {""Passing Sck"", Int64.Type}, {""Passing SckY"", type number}, {""Passing Rate"", type number}, {""Rushing Att"", Int64.Type}, {""Rushing Yds"", type number}, {""Rushing Avg"", type number}, {""Rushin" & _
        "g TD"", Int64.Type}, {""Fumbles FUM"", Int64.Type}, {""Fumbles Lost"", Int64.Type}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 1"";Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 1]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = name_table
        .Refresh BackgroundQuery:=False
    End With
    ActiveWorkbook.Queries("Table 1").Name = name_table
End Sub
Private Sub add_per_game_table(ByVal name_table)

    Dim last_used_row As Integer
    Dim last_used_col As Integer
    Dim last_table_col As Integer
    Dim import_row As Integer
    Dim col_count_1 As Integer
    Dim col_count_2 As Integer
    Dim row_counter_1 As Integer
    Dim end_import_table_row As Integer
    Dim current_row As Integer
    Dim stats_start_col As Integer
    Dim gp_col As Integer
    Dim num_import_rows As Integer
    Dim col_header As Variant

    With Sheets(name_table)
        last_used_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        last_used_col = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        .Cells(1, last_used_col + 1).Value = "FPs"
               
        col_header = .Range(.Cells(1, 1), .Cells(1, last_used_col))
        import_row = last_used_row + 2
        num_import_rows = last_used_row - 2
        end_import_table_row = num_import_rows + import_row

        .ListObjects.Add(xlSrcRange, .Range(.Cells(import_row, 1), .Cells(end_import_table_row - 1, last_used_col + 1)), , xlNo).Name = _
            name_table & "_per_game"

        .Cells(import_row, last_used_col + 1).Value = "FPs"
        .Range(.Cells(import_row, 1), .Cells(import_row, last_used_col)) = col_header
        current_row = 1
        For row_counter_1 = import_row + 1 To end_import_table_row
            current_row = current_row + 1
            For col_count_1 = 1 To last_used_col + 1
                If .Cells(current_row, col_count_1).Value <> "" Or .Cells(current_row, col_count_1).Value = "0" Then
                    If .Cells(1, col_count_1) = "G" Then
                        gp_col = col_count_1
                        .Cells(row_counter_1, col_count_1) = .Cells(current_row, col_count_1)
                    ElseIf .Cells(1, col_count_1) = "Passing Comp" Or _
                        .Cells(1, col_count_1) = "Passing Att" Or _
                        .Cells(1, col_count_1) = "Passing Yds" Or _
                        .Cells(1, col_count_1) = "Passing TD" Or _
                        .Cells(1, col_count_1) = "Passing Int" Or _
                        .Cells(1, col_count_1) = "Passing Sck" Or _
                        .Cells(1, col_count_1) = "Passing SckY" Or _
                        .Cells(1, col_count_1) = "Rushing Att" Or _
                        .Cells(1, col_count_1) = "Rushing Yds" Or _
                        .Cells(1, col_count_1) = "Rushing TD" Or _
                        .Cells(1, col_count_1) = "Fumbles FUM" Or _
                        .Cells(1, col_count_1) = "Fumbles Lost" Then
    
                        If .Cells(1, col_count_1) = "Passing Comp" Then
                            stats_start_col = col_count_1
                        End If
                        .Cells(row_counter_1, col_count_1) = .Cells(current_row, col_count_1) / .Cells(current_row, gp_col)
                    ElseIf col_count_1 > last_used_col Then
                        .Cells(current_row, col_count_1).FormulaR1C1 = _
                            "=(0.04*RC[-13])+(6*RC[-11]) -(2*RC[-10])+(0.1*RC[-5])+(6*RC[-3])-(2*RC[-1])"
                        .Cells(row_counter_1, col_count_1).FormulaR1C1 = _
                            "=(0.04*RC[-13])+(6*RC[-11]) -(2*RC[-10])+(0.1*RC[-5])+(6*RC[-3])-(2*RC[-1])"
                    Else
                        .Cells(row_counter_1, col_count_1) = .Cells(current_row, col_count_1)
                    End If
                ElseIf col_count_1 > last_used_col Then
                    .Cells(current_row, col_count_1).FormulaR1C1 = _
                        "=(0.04*RC[-13])+(6*RC[-11]) -(2*RC[-10])+(0.1*RC[-5])+(6*RC[-3])-(2*RC[-1])"
                    .Cells(row_counter_1, col_count_1).FormulaR1C1 = _
                        "=(0.04*RC[-13])+(6*RC[-11]) -(2*RC[-10])+(0.1*RC[-5])+(6*RC[-3])-(2*RC[-1])"
                Else
                    .Cells(row_counter_1, col_count_1) = .Cells(current_row, col_count_1)
                End If
            Next
        Next
        .Range(.Cells(import_row + 1, stats_start_col), .Cells(end_import_table_row, last_used_col)).NumberFormat = "0.0"
        add_16_game_table name_table, num_import_rows, import_row + 1
    End With
End Sub
Private Sub add_16_game_table(ByVal name_table, ByVal num_import_rows, ByVal data_start_row)

    Dim last_used_row As Integer
    Dim last_used_col As Integer
    Dim last_table_col As Integer
    Dim import_row As Integer
    Dim col_count_1 As Integer
    Dim col_count_2 As Integer
    Dim row_counter_1 As Integer
    Dim end_import_table_row As Integer
    Dim current_row As Integer
    Dim stats_start_col As Integer
    Dim col_header As Variant

    With Sheets(name_table)
        last_used_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        last_used_col = .Cells(1, .Columns.Count).End(xlToLeft).Column
        col_header = .Range(.Cells(1, 1), .Cells(1, last_used_col))
        import_row = last_used_row + 2
        end_import_table_row = import_row + num_import_rows

        .ListObjects.Add(xlSrcRange, .Range(.Cells(import_row, 1), .Cells(end_import_table_row - 1, last_used_col)), , xlNo).Name = _
            name_table & "_16_game"

        .Cells(import_row, last_used_col).Value = "FPs"
        .Range(.Cells(import_row, 1), .Cells(import_row, last_used_col)) = col_header
        current_row = data_start_row - 1
        For row_counter_1 = import_row + 1 To end_import_table_row
            current_row = current_row + 1
            For col_count_1 = 1 To last_used_col
                If .Cells(current_row, col_count_1).Value <> "" Then
                    If .Cells(1, col_count_1) = "Passing Comp" Or _
                        .Cells(1, col_count_1) = "Passing Att" Or _
                        .Cells(1, col_count_1) = "Passing Yds" Or _
                        .Cells(1, col_count_1) = "Passing TD" Or _
                        .Cells(1, col_count_1) = "Passing Int" Or _
                        .Cells(1, col_count_1) = "Passing Sck" Or _
                        .Cells(1, col_count_1) = "Passing SckY" Or _
                        .Cells(1, col_count_1) = "Rushing Att" Or _
                        .Cells(1, col_count_1) = "Rushing Yds" Or _
                        .Cells(1, col_count_1) = "Rushing TD" Or _
                        .Cells(1, col_count_1) = "Fumbles FUM" Or _
                        .Cells(1, col_count_1) = "Fumbles Lost" Then
    
                        If .Cells(1, col_count_1) = "Passing Comp" Then
                            stats_start_col = col_count_1
                        End If
                        .Cells(row_counter_1, col_count_1) = .Cells(current_row, col_count_1) * 16
                    ElseIf col_count_1 = last_used_col Then
                        .Cells(row_counter_1, col_count_1).FormulaR1C1 = _
                            "=(0.04*RC[-13])+(6*RC[-11]) -(2*RC[-10])+(0.1*RC[-5])+(6*RC[-3])-(2*RC[-1])"
                    Else
                        .Cells(row_counter_1, col_count_1) = .Cells(current_row, col_count_1)
                    End If
                ElseIf col_count_1 = last_used_col Then
                    .Cells(row_counter_1, col_count_1).FormulaR1C1 = _
                        "=(0.04*RC[-13])+(6*RC[-11]) -(2*RC[-10])+(0.1*RC[-5])+(6*RC[-3])-(2*RC[-1])"
                Else
                    .Cells(row_counter_1, col_count_1) = .Cells(current_row, col_count_1)
                End If
            Next
        Next
        .Range(.Cells(import_row + 1, stats_start_col), .Cells(end_import_table_row, last_used_col)).NumberFormat = "0.0"
    End With
End Sub

