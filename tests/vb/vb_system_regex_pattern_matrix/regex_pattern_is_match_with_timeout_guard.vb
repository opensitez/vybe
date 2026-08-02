' vybe-test: vb/vb_system_regex_pattern_matrix/regex_pattern_is_match_with_timeout_guard
' origin: languages/vb/tests/vb/test_vb_system_regex_pattern_matrix.rs

Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim value As String = "abc123"
        Dim ok As Boolean = Regex.IsMatch(value, "[a-z]+\d+", RegexOptions.None)
        Console.WriteLine(ok)

        Dim hasTimeout As Boolean = False
        Try
            Dim hit As Boolean = Regex.IsMatch("abc", "a.*b", RegexOptions.None, TimeSpan.FromMilliseconds(100))
            Console.WriteLine(hit)
        Catch ex As RegexMatchTimeoutException
            hasTimeout = True
            Console.WriteLine("to")
        End Try

        If hasTimeout Then
            Console.WriteLine(True)
        Else
            Console.WriteLine(ok)
        End If
    End Sub
End Module
