'https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/

Private Declare PtrSafe Function GetKeyState Lib "User32" (ByVal vKey As Integer) As Integer 'remove PtrSafe for 32-bit

Const ALT_KEY = 18

' Remember:
' Rows and columns start at index one (1)


' Location of the fretboards top left corner:
Const fret_offset_row_c = 4
Const fret_offset_column_c = 2


' Location of the printed notes
Const tabs_start_row_c = 14
Const tabs_end_row_c = tabs_start_row_c + 5     ' Add (5) to cover all six strings
Const fret_range As String = "B14:AG19"


' We use this one to decide where to write out the tabs (or clicks)
Const fret_to_tab_distance_c = tabs_start_row_c - fret_offset_row_c


' Clear button location
Const clear_btn_offset_column_c = 33


' Event: selection
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  write_frets (Cells(Target.Row, Target.Column))
End Sub


' Event: Double click which can overwrite following below row(s)
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  copy_fretboard
  Cancel = True
End Sub



Private Sub write_frets(Target As Range)

  ' To disable research tab from opening when pressing Alt key and clicking one cell
  Application.CommandBars("Research").Enabled = False


  ' If we clicked inside the fret board, i.e.,
  ' fret 0 to 22 and the blank column
  If Target.Row >= fret_offset_row_c And _
     Target.Row <= fret_offset_row_c + 5 And _
     Target.Column >= fret_offset_column_c And _
     Target.Column <= fret_offset_column_c + 23 _
  Then

    ' to check which column to write to
    write_column = get_rightmost_column(tabs_start_row_c, tabs_end_row_c) + 1

    If GetKeyState(ALT_KEY) < 0 Then
      write_column = write_column - 1
    End If

    ' make sure holding ALT doesn't change the first column
    If write_column = 1 Then
      write_column = 2
    End If

    write_row = Target.Row + fret_to_tab_distance_c

    clicked_fret_number = Target.Column - fret_offset_column_c

    If clicked_fret_number = 23 Then
      Cells(write_row, write_column) = " "
    Else
      Cells(write_row, write_column) = Target.Column - fret_offset_column_c
    End If

  End If


  ' If clear button was clicked
  If Target.Column = clear_btn_offset_column_c And _
     Target.Row >= fret_offset_row_c And _
     Target.Row <= fret_offset_row_c + 5 _
  Then

    Rows(tabs_start_row_c & ":" & tabs_end_row_c).Cells.ClearContents

  End If

End Sub



Private Function get_rightmost_column(start_row As Variant, end_row As Variant)

  max_column = 0

  For i = start_row To end_row

    max_in_row = Cells(i, Cells.Columns.Count).End(xlToLeft).Column

    If max_in_row > max_column Then
      max_column = max_in_row
    End If

  Next i

  get_rightmost_column = max_column

End Function



Sub copy_fretboard()

  max_column = get_rightmost_column(tabs_start_row_c, tabs_end_row_c) + 1

  If max_column <= 2 Then Exit Sub ' first column is not copied

  Range(Cells(tabs_start_row_c, max_column), _
        Cells(tabs_end_row_c, max_column)).Value = Range(Cells(tabs_start_row_c, max_column - 1), _
                                                         Cells(tabs_end_row_c, max_column - 1)).Value

End Sub

