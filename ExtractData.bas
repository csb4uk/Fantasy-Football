Attribute VB_Name = "ExtractData"
Option Explicit
Option Base 1
Public Sub extract_data()

    Dim summary_wks As Worksheet

    Dim starter_rows As New Collection
    Dim bench_rows As New Collection
    Dim outscore_coll As New Collection
    Dim max_coll As New Collection

    Dim header_row As Integer
    Dim last_column As Integer

    Set summary_wks = Worksheets("Summary")
    With summary_wks.UsedRange
        last_column = .Columns(.Columns.Count).Column
    End With

    'Id the rows that contain bench players
    id_bench_rows summary_wks, starter_rows, bench_rows, header_row
    'Id number of players that outscored starters
    id_outscore summary_wks, last_column, header_row, bench_rows, starter_rows, outscore_coll, max_coll
    'Write outscored collection to stats sheet
    id_write_rows outscore_coll, max_coll
End Sub

Private Sub id_bench_rows(ByVal summary_wks, ByRef starter_rows, ByRef bench_rows, ByRef header_row)

    Dim start_row_summary As Integer
    Dim last_row_summary As Integer
    Dim row_counter_summary As Integer

    start_row_summary = 1
    With summary_wks
        last_row_summary = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    For row_counter_summary = start_row_summary To last_row_summary
        If InStr(1, summary_wks.Range("A" & row_counter_summary).Value, "Bench") > 0 Then
            bench_rows.Add row_counter_summary
        Else
            Select Case summary_wks.Range("A" & row_counter_summary).Value
                Case "QB", "RB1", "RB2", "WR1", "WR2", "TE", "FLEX", "D/ST"
                    starter_rows.Add row_counter_summary
                Case "SLOT"
                    header_row = row_counter_summary
                Case Else
            End Select
        End If
    Next
End Sub

Private Sub id_outscore(ByVal summary_wks, ByVal last_column, ByVal header_row, ByVal bench_rows, ByVal starter_rows, ByRef outscore_coll, ByRef max_coll)

    Dim column_counter As Integer
    Dim outscore_counter As Integer

    Dim bench_player_points As Double
    Dim starter_player_points As Double
    
    Dim bench_row As Variant
    Dim starter_row As Variant

    Dim bench_player_pos As String
    Dim slot_position As String
    Dim starter_player_pos As String


    For column_counter = 1 To last_column
        If summary_wks.Cells(header_row, column_counter) = "Points" Then
            For Each bench_row In bench_rows
                If summary_wks.Cells(bench_row, column_counter - 1) <> "" Then
                    bench_player_points = summary_wks.Cells(bench_row, column_counter)
                    bench_player_pos = summary_wks.Cells(bench_row, column_counter - 1)
                    If bench_player_points <> 0 Then
                        For Each starter_row In starter_rows
                            slot_position = summary_wks.Cells(starter_row, 1).Value
                            starter_player_points = summary_wks.Cells(starter_row, column_counter).Value
                            starter_player_pos = summary_wks.Cells(starter_row, column_counter - 1).Value
                            If slot_position <> "FLEX" Then
                                If bench_player_pos = starter_player_pos And bench_player_points > starter_player_points Then
                                    outscore_counter = outscore_counter + 1
                                    GoTo NextBenchBreak
                                End If
                            Else
                                If slot_position <> "QB" Or slot_position <> "D/ST" Then
                                    If bench_player_points > starter_player_points Then
                                        outscore_counter = outscore_counter + 1
                                        GoTo NextBenchBreak
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
NextBenchBreak:
            Next
        outscore_coll.Add outscore_counter
        outscore_counter = 0

        'ID max score
        id_max_score column_counter, summary_wks, header_row, max_coll
        End If
    Next
End Sub

Private Sub id_max_score(ByVal current_column, ByVal summary_wks, ByVal header_row, ByRef max_coll)

    Dim data_array As Variant

    Dim last_row_summary As Integer
    Dim array_counter As Integer

    Dim previous_max As Double
    Dim qb_max As Double
    Dim rb1_max As Double
    Dim rb2_max As Double
    Dim wr1_max As Double
    Dim wr2_max As Double
    Dim te_max As Double
    Dim dst_max As Double
    Dim flex_max As Double
    Dim points As Double
    Dim score_max As Double

    Dim position As String

    qb_max = -10
    rb1_max = -10
    rb2_max = -10
    wr1_max = -10
    wr2_max = -10
    te_max = -10
    dst_max = -10
    flex_max = -10

    With summary_wks
        last_row_summary = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    data_array = summary_wks.Range(summary_wks.Cells(header_row + 1, current_column - 1), summary_wks.Cells(last_row_summary, current_column))

    For array_counter = 1 To UBound(data_array, 1)
        If data_array(array_counter, 1) <> "" And data_array(array_counter, 2) <> "" Then
            position = data_array(array_counter, 1)
            points = data_array(array_counter, 2)
RecalcBreak:
            Select Case position
                Case "QB"
                    If points > qb_max Then
                        qb_max = points
                    End If
                Case "RB"
                    If points > rb1_max Then
                        previous_max = rb1_max
                        rb1_max = points
                        points = previous_max
                        GoTo RecalcBreak
                    ElseIf points > rb2_max Then
                        previous_max = rb2_max
                        rb2_max = points
                        points = previous_max
                        GoTo RecalcBreak
                    ElseIf points > flex_max Then
                        previous_max = flex_max
                        flex_max = points
                        points = previous_max
                        GoTo RecalcBreak
                    End If
                Case "WR"
                    If points > wr1_max Then
                        previous_max = wr1_max
                        wr1_max = points
                        points = previous_max
                        GoTo RecalcBreak
                    ElseIf points > wr2_max Then
                        previous_max = wr2_max
                        wr2_max = points
                        points = previous_max
                        GoTo RecalcBreak
                    ElseIf points > flex_max Then
                        previous_max = flex_max
                        flex_max = points
                        points = previous_max
                        GoTo RecalcBreak
                    End If
                Case "TE"
                    If points > te_max Then
                        previous_max = te_max
                        te_max = points
                        points = previous_max
                        GoTo RecalcBreak
                    ElseIf points > flex_max Then
                        previous_max = flex_max
                        flex_max = points
                        points = previous_max
                        GoTo RecalcBreak
                    End If
                Case "D/ST"
                    If points > dst_max Then
                        dst_max = points
                    End If
            End Select
        End If
    Next
    score_max = qb_max + rb1_max + rb2_max + wr1_max + wr2_max + te_max + dst_max + flex_max
    max_coll.Add score_max
End Sub

Private Sub id_write_rows(ByVal outscore_coll, ByVal max_coll)
    Dim outscored_write_row As Integer
    Dim max_write_row As Integer
    Dim stat_row_count As Integer
    Dim stat_last_row As Integer
    Dim stat_column_count As Integer
    Dim stat_last_column As Integer
    Dim collection_counter As Integer

    Dim column_coll As New Collection
    
    Dim stats_wks As Worksheet

    Set stats_wks = Sheets("Stats")
    With stats_wks
        stat_last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    With stats_wks.UsedRange
        stat_last_column = .Columns(.Columns.Count).Column
    End With
    For stat_row_count = 1 To stat_last_row
        If stats_wks.Cells(stat_row_count, 1) = "Bench Players Outscored Starters" Then
            outscored_write_row = stat_row_count
            Exit For
        End If
    Next
    For stat_row_count = 1 To stat_last_row
        If stats_wks.Cells(stat_row_count, 1) = "Max Score" Then
            max_write_row = stat_row_count
            Exit For
        End If
    Next
    For stat_column_count = 1 To stat_last_column
        If InStr(1, stats_wks.Cells(1, stat_column_count).Value, "Week") > 0 Then
            column_coll.Add stat_column_count
        End If
    Next
    For collection_counter = 1 To column_coll.Count
        stats_wks.Cells(outscored_write_row, column_coll(collection_counter)).Value = outscore_coll(collection_counter)
        stats_wks.Cells(max_write_row, column_coll(collection_counter)).Value = max_coll(collection_counter)
    Next
End Sub

