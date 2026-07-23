use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Text.RegularExpressions.Regex & Named Groups
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_regex_named_group_extraction() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "User: Alice, Age: 30"
        Dim pattern = "User: (?<name>\w+), Age: (?<age>\d+)"
        Dim match = Regex.Match(input, pattern)
        Console.WriteLine(match.Groups("name").Value & "|" & match.Groups("age").Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice|30"]);
}

#[test]
fn test_vb_regex_matches_multiple_occurrences() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "cat, dog, bird"
        Dim matches = Regex.Matches(input, "\b\w+\b")
        Console.WriteLine(matches.Count & ":" & matches(0).Value & "," & matches(1).Value & "," & matches(2).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3:cat,dog,bird"]);
}

#[test]
fn test_vb_regex_options_ignore_case() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim isMatch = Regex.IsMatch("VISUALBASIC", "visualbasic", RegexOptions.IgnoreCase)
        Console.WriteLine(isMatch)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_regex_options_singleline_dot_matches_newline() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "Start" & vbCrLf & "End"
        Dim match = Regex.Match(input, "Start.*End", RegexOptions.Singleline)
        Console.WriteLine(match.Success)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_regex_options_multiline_anchor_matching() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "Line1" & vbLf & "Line2" & vbLf & "Line3"
        Dim matches = Regex.Matches(input, "^Line\d", RegexOptions.Multiline)
        Console.WriteLine(matches.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_regex_split_string_by_pattern() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "one1two2three3four"
        Dim parts = Regex.Split(input, "\d")
        Console.WriteLine(String.Join(",", parts))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["one,two,three,four"]);
}

#[test]
fn test_vb_regex_escape_and_unescape() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim raw = "c:\folder\file.txt?val=1*2"
        Dim escaped = Regex.Escape(raw)
        Dim unescaped = Regex.Unescape(escaped)
        Console.WriteLine((raw = unescaped) & "|" & escaped.Contains("\?"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_regex_group_success_and_index_length() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim match = Regex.Match("ABC 123 XYZ", "\d+")
        Console.WriteLine(match.Success & "|" & match.Index & "|" & match.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|4|3"]);
}

#[test]
fn test_vb_regex_compiled_option_performance() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim regex As New Regex("\d+", RegexOptions.Compiled)
        Dim m = regex.Match("Count: 99")
        Console.WriteLine(m.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}

#[test]
fn test_vb_regex_get_group_names() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim regex As New Regex("(?<first>\w+)\s+(?<last>\w+)")
        Dim groupNames = regex.GetGroupNames()
        Console.WriteLine(String.Join(",", groupNames))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,first,last"]);
}

#[test]
fn test_vb_regex_group_name_from_number() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim regex As New Regex("(?<area>\d{3})-(?<num>\d{4})")
        Console.WriteLine(regex.GroupNameFromNumber(1) & "|" & regex.GroupNameFromNumber(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["area|num"]);
}

#[test]
fn test_vb_regex_group_number_from_name() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim regex As New Regex("(?<area>\d{3})-(?<num>\d{4})")
        Console.WriteLine(regex.GroupNumberFromName("area") & "|" & regex.GroupNumberFromName("num"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|2"]);
}

#[test]
fn test_vb_regex_captures_collection_multiple_captures() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim match = Regex.Match("123 456 789", "(\d+\s*)+")
        Dim captures = match.Groups(1).Captures
        Console.WriteLine(captures.Count & ":" & captures(0).Value.Trim() & "," & captures(1).Value.Trim() & "," & captures(2).Value.Trim())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3:123,456,789"]);
}

#[test]
fn test_vb_regex_non_capturing_group() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim match = Regex.Match("Value: 42", "(?:Value:\s*)(\d+)")
        Console.WriteLine(match.Groups.Count & "|" & match.Groups(1).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|42"]);
}

#[test]
fn test_vb_regex_positive_lookahead() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        ' Match word only if followed by " USD"
        Dim match = Regex.Match("100 USD, 200 EUR", "\d+(?=\s*USD)")
        Console.WriteLine(match.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_regex_negative_lookahead() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        ' Match digits not followed by " USD"
        Dim matches = Regex.Matches("100 USD, 200 EUR", "\d+\b(?!\s*USD)")
        Console.WriteLine(matches(0).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["200"]);
}

#[test]
fn test_vb_regex_timeout_cancellation() {
    let src = r#"
Imports System
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim regex As New Regex("(a+)+$", RegexOptions.None, TimeSpan.FromMilliseconds(100))
        Try
            regex.IsMatch("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaX")
        Catch ex As RegexMatchTimeoutException
            Console.WriteLine("RegexMatchTimeoutException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["RegexMatchTimeoutException Caught"]);
}

#[test]
fn test_vb_regex_backreference_matching() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        ' Match repeated word like "the the"
        Dim match = Regex.Match("Is the the end?", "\b(\w+)\s+\1\b")
        Console.WriteLine(match.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["the the"]);
}

#[test]
fn test_vb_regex_right_to_left_option() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim match = Regex.Match("100 200 300", "\d+", RegexOptions.RightToLeft)
        Console.WriteLine(match.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["300"]);
}

#[test]
fn test_vb_regex_explicit_capture_option() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        ' ExplicitCapture means un-named (group) is NOT captured!
        Dim regex As New Regex("(unnamed) (?<named>val)", RegexOptions.ExplicitCapture)
        Dim m = regex.Match("unnamed val")
        Console.WriteLine(m.Groups.Count & "|" & m.Groups("named").Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|val"]);
}
