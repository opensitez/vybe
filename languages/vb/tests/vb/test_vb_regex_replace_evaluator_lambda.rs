use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Regex.Replace, MatchEvaluator Lambdas & Substitution Patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_regex_replace_match_evaluator_lambda() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "cat 10 dog 20"
        Dim result = Regex.Replace(input, "\d+", Function(m) (Integer.Parse(m.Value) * 2).ToString())
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["cat 20 dog 40"]);
}

#[test]
fn test_vb_regex_replace_named_groups_substitution() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "John Smith"
        Dim result = Regex.Replace(input, "(?<first>\w+)\s+(?<last>\w+)", "${last}, ${first}")
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Smith, John"]);
}

#[test]
fn test_vb_regex_replace_numbered_groups_substitution() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "2025-12-31"
        Dim result = Regex.Replace(input, "(\d{4})-(\d{2})-(\d{2})", "$2/$3/$1")
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12/31/2025"]);
}

#[test]
fn test_vb_regex_replace_with_count_limit() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "a b a b a"
        Dim result = Regex.Replace(input, "a", "X", RegexOptions.None, TimeSpan.FromSeconds(1))
        Dim resultLimit = New Regex("a").Replace(input, "X", 2)
        Console.WriteLine(resultLimit)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X b X b a"]);
}

#[test]
fn test_vb_regex_replace_entire_match_dollarsign_zero() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "100 200"
        Dim result = Regex.Replace(input, "\d+", "[$0]")
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[100] [200]"]);
}

#[test]
fn test_vb_regex_replace_text_before_and_after_match() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "AAA MATCH BBB"
        Dim result = Regex.Replace(input, "MATCH", "[$`|$']")
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AAA [AAA | BBB] BBB"]);
}

#[test]
fn test_vb_regex_replace_last_captured_group() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "A-B-C"
        Dim result = Regex.Replace(input, "([A-Z]-)+", "$+")
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B-C"]);
}

#[test]
fn test_vb_regex_replace_evaluator_uppercase_conversion() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "hello world from vb"
        Dim result = Regex.Replace(input, "\b\w+\b", Function(m) m.Value.ToUpper())
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["HELLO WORLD FROM VB"]);
}

#[test]
fn test_vb_regex_replace_evaluator_with_index_tracking() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "apple banana cherry"
        Dim index = 1
        Dim result = Regex.Replace(input, "\b\w+\b", Function(m)
            Dim res = index & "." & m.Value
            index += 1
            Return res
        End Function)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.apple 2.banana 3.cherry"]);
}

#[test]
fn test_vb_regex_replace_literal_dollar_sign() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "Cost: 100"
        ' $$ produces literal $ in replacement string!
        Dim result = Regex.Replace(input, "\d+", "$$$0")
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Cost: $100"]);
}

#[test]
fn test_vb_regex_replace_start_at_offset() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "tag tag tag"
        Dim regex As New Regex("tag")
        Dim result = regex.Replace(input, "KEY", -1, 4) ' Start replacing at index 4
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["tag KEY KEY"]);
}

#[test]
fn test_vb_regex_replace_html_sanitization_simulation() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "<b>Hello</b> <i>World</i>"
        Dim clean = Regex.Replace(input, "<[^>]+>", "")
        Console.WriteLine(clean)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_regex_replace_whitespace_normalization() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "  Too    much   space   "
        Dim clean = Regex.Replace(input.Trim(), "\s+", " ")
        Console.WriteLine(clean)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Too much space"]);
}

#[test]
fn test_vb_regex_replace_mask_sensitive_data() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim cc = "1234-5678-9012-3456"
        Dim masked = Regex.Replace(cc, "\d{4}-\d{4}-\d{4}-", "****-****-****-")
        Console.WriteLine(masked)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["****-****-****-3456"]);
}

#[test]
fn test_vb_regex_replace_camel_to_snake_case() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim camel = "FirstName"
        Dim snake = Regex.Replace(camel, "(?<!^)([A-Z])", "_$1").ToLower()
        Console.WriteLine(snake)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["first_name"]);
}

#[test]
fn test_vb_regex_replace_evaluator_conditional_replacement() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "10 25 50 100"
        Dim result = Regex.Replace(input, "\d+", Function(m)
            Dim val = Integer.Parse(m.Value)
            If val > 30 Then Return "HIGH" Else Return "LOW"
        End Function)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["LOW LOW HIGH HIGH"]);
}

#[test]
fn test_vb_regex_replace_non_matching_returns_original() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim input = "NoDigitsHere"
        Dim result = Regex.Replace(input, "\d+", "X")
        Console.WriteLine(Object.ReferenceEquals(input, result) OrElse input = result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_regex_replace_empty_input() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim result = Regex.Replace("", "\d+", "X")
        Console.WriteLine(result = "")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_regex_replace_compiled_instance() {
    let src = r#"
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim regex As New Regex("\s+", RegexOptions.Compiled)
        Dim result = regex.Replace("A B  C   D", "_")
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A_B_C_D"]);
}

#[test]
fn test_vb_regex_replace_evaluator_exception_propagation() {
    let src = r#"
Imports System
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Try
            Regex.Replace("test", "test", Function(m)
                Throw New InvalidOperationException("Evaluator Error")
            End Function)
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Evaluator Error"]);
}
