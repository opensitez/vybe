use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Regex Replacement & Named Match Groups
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_regex_named_group_match() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim pattern As String = "(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})"
        Dim m As Match = Regex.Match("2026-07-21", pattern)
        Console.WriteLine(m.Groups("year").Value)
        Console.WriteLine(m.Groups("month").Value)
        Console.WriteLine(m.Groups("day").Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2026", "07", "21"]);
}

#[test]
fn test_vb_regex_replace_backreference() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input As String = "John Smith"
        Dim pattern As String = "(\w+)\s+(\w+)"
        Dim output As String = Regex.Replace(input, pattern, "$2, $1")
        Console.WriteLine(output)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Smith, John"]);
}

#[test]
fn test_vb_regex_replace_evaluator_delegate() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input As String = "10 20 30"
        Dim evaluator As MatchEvaluator = Function(m As Match) (Integer.Parse(m.Value) * 2).ToString()
        Dim output As String = Regex.Replace(input, "\d+", evaluator)
        Console.WriteLine(output)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20 40 60"]);
}

#[test]
fn test_vb_regex_matches_collection() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim matches As MatchCollection = Regex.Matches("cat mat sat", "\w+at")
        Console.WriteLine(matches.Count)
        For Each m As Match In matches
            Console.WriteLine(m.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "cat", "mat", "sat"]);
}

#[test]
fn test_vb_regex_split_by_digits() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim parts As String() = Regex.Split("one1two2three", "\d")
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(0))
        Console.WriteLine(parts(1))
        Console.WriteLine(parts(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "one", "two", "three"]);
}

#[test]
fn test_vb_regex_options_multiline_ignore_case() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim text As String = "hello" & vbLf & "WORLD"
        Dim options As RegexOptions = RegexOptions.IgnoreCase Or RegexOptions.Multiline
        Dim m As MatchCollection = Regex.Matches(text, "^[a-z]+$", options)
        Console.WriteLine(m.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_regex_group_captures_collection() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim m As Match = Regex.Match("123-456", "(\d+)-(\d+)")
        Console.WriteLine(m.Groups.Count)
        Console.WriteLine(m.Groups(1).Value)
        Console.WriteLine(m.Groups(2).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "123", "456"]);
}

#[test]
fn test_vb_regex_is_match_timeout() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim isMatch As Boolean = Regex.IsMatch("abc", "^a", RegexOptions.None, TimeSpan.FromSeconds(1))
        Console.WriteLine(isMatch)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_regex_escape_unescape() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim escaped As String = Regex.Escape("a+b*c?")
        Dim unescaped As String = Regex.Unescape(escaped)
        Console.WriteLine(escaped)
        Console.WriteLine(unescaped)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec![r"a\+b\*c\?", "a+b*c?"]);
}

#[test]
fn test_vb_regex_compiled_instance() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim re As New Regex("\d+", RegexOptions.Compiled)
        Dim m As Match = re.Match("Item 100")
        Console.WriteLine(m.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_regex_group_success_and_index() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim m As Match = Regex.Match("abc 123 xyz", "(?<digits>\d+)")
        Dim g As Group = m.Groups("digits")
        Console.WriteLine(g.Success)
        Console.WriteLine(g.Index)
        Console.WriteLine(g.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "4", "3"]);
}

#[test]
fn test_vb_regex_replace_count_limit() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim re As New Regex("x")
        Dim res As String = re.Replace("x x x x", "y", 2)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["y y x x"]);
}

#[test]
fn test_vb_regex_lookahead_assertion() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim m As Match = Regex.Match("100USD 200EUR", "\d+(?=USD)")
        Console.WriteLine(m.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_regex_lookbehind_assertion() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim m As Match = Regex.Match("$100 €200", "(?<=\$)\d+")
        Console.WriteLine(m.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_regex_non_capturing_group() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim m As Match = Regex.Match("FooBar", "(?:Foo)(Bar)")
        Console.WriteLine(m.Groups.Count)
        Console.WriteLine(m.Groups(1).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "Bar"]);
}
