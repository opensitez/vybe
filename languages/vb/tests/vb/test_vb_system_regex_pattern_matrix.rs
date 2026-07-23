use super::helpers::run_vb;

#[test]
fn regex_pattern_simple_match_capture() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim pattern As String = "^(?<first>\w+)-(?<second>\d+)$"
        Dim value As String = "item-123"
        Dim m As Match = Regex.Match(value, pattern)

        Console.WriteLine(m.Success)
        Console.WriteLine(m.Groups("first").Value)
        Console.WriteLine(m.Groups("second").Value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "item", "123"]);
}

#[test]
fn regex_pattern_ignore_case_and_multiline() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim input As String = "Line1\nline2\nLINE3"
        Dim matches As MatchCollection = Regex.Matches(input, "line", RegexOptions.IgnoreCase Or RegexOptions.Multiline)

        Console.WriteLine(matches.Count)
        Console.WriteLine(matches(0).Value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "Line"]);
}

#[test]
fn regex_pattern_replace_and_split() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim source As String = "a1 b2 c3"
        Dim clean As String = Regex.Replace(source, "\d", "")
        Dim parts As String() = Regex.Split("a,b;;c", "[;,]+")

        Console.WriteLine(clean = "a b c")
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(0))
        Console.WriteLine(parts(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "3", "a", "c"]);
}

#[test]
fn regex_pattern_is_match_with_timeout_guard() {
    let out = run_vb(
        r#"
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
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}
