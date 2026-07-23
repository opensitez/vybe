use super::helpers::run_vb;

#[test]
fn regex_named_group_capture() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim m As Match = Regex.Match("date=2024-06-15", "(?<year>\d{4})-(?<month>\d{2})")
        Console.WriteLine(m.Groups("year").Value)
        Console.WriteLine(m.Groups("month").Value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2024", "06"]);
}

#[test]
fn regex_matches_returns_all_occurrences() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim matches As MatchCollection = Regex.Matches("a1 b2 c3", "\d")
        Console.WriteLine(matches.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3"]);
}

#[test]
fn regex_replace_with_match_evaluator_transforms_matches() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim result As String = Regex.Replace(
            "a1b2c3",
            "\d",
            Function(m As Match) (CInt(m.Value) * 2).ToString()
        )
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["a2b4c6"]);
}

#[test]
fn regex_anchored_pattern_rejects_midstring_match() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Console.WriteLine(Regex.IsMatch("abc", "^\d+$"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn regex_character_class_matches_expected_char() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim m As Match = Regex.Match("hello", "[aeiou]")
        Console.WriteLine(m.Value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["e"]);
}

#[test]
fn regex_quantifier_plus_requires_at_least_one_digit() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Console.WriteLine(Regex.IsMatch("007", "\d+"))
        Console.WriteLine(Regex.IsMatch("", "\d+"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn regex_multiline_affects_caret_matches() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim matches As MatchCollection = Regex.Matches(
            "start\nnew line",
            "^[a-z]",
            RegexOptions.Multiline
        )
        Console.WriteLine(matches.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn regex_split_on_word_boundaries() {
    let out = run_vb(
        r#"
Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim parts As String() = Regex.Split("a,b; c", "[,;]\\s*")
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(0))
        Console.WriteLine(parts(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "a", "c"]);
}
