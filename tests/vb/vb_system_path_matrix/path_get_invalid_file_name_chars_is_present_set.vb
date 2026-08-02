' vybe-test: vb/vb_system_path_matrix/path_get_invalid_file_name_chars_is_present_set
' origin: languages/vb/tests/vb/test_vb_system_path_matrix.rs

Imports System.IO

Module M
    Sub Main()
        Dim bad() As Char = Path.GetInvalidFileNameChars()
        Dim badCount As Integer = bad.Length

        Dim ok As Boolean = False
        For i As Integer = 0 To bad.Length - 1
            If bad(i) <> ""C Then
                ok = True
                Exit For
            End If
        Next

        Console.WriteLine(badCount >= 0)
        Console.WriteLine(ok)
    End Sub
End Module
