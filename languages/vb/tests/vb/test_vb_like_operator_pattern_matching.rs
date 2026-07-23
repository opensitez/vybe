use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Like Operator Pattern Matching & Wildcards
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_like_operator_asterisk_wildcard() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("VisualBasic" Like "Visual*") & "|" & ("VisualBasic" Like "*Basic") & "|" & ("VisualBasic" Like "*al*"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_like_operator_question_mark_single_char() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("cat" Like "c?t") & "|" & ("cart" Like "c?t"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_like_operator_hash_single_digit() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("A5B" Like "A#B") & "|" & ("AXB" Like "A#B"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_like_operator_char_list_brackets() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("bat" Like "[bcr]at") & "|" & ("fat" Like "[bcr]at"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_like_operator_char_range_brackets() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("m" Like "[a-z]") & "|" & ("5" Like "[a-z]"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_like_operator_negated_char_list_exclamation() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("bat" Like "[!cr]at") & "|" & ("cat" Like "[!cr]at"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_like_operator_case_sensitivity_option_compare_binary() {
    let src = r#"
Option Compare Binary

Module Program
    Sub Main()
        Console.WriteLine("abc" Like "ABC")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_like_operator_case_insensitivity_option_compare_text() {
    let src = r#"
Option Compare Text

Module Program
    Sub Main()
        Console.WriteLine("abc" Like "ABC")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_like_operator_escaped_special_chars_in_brackets() {
    let src = r#"
Module Program
    Sub Main()
        ' In Like operator, special chars like *, ?, # can be matched by putting them in brackets!
        Console.WriteLine(("100#" Like "100[#]") & "|" & ("100*" Like "100[*]"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_like_operator_multiple_wildcards_combined() {
    let src = r#"
Module Program
    Sub Main()
        Dim pattern = "User_#_??_[a-z]*"
        Dim str = "User_1_AB_data.txt"
        Console.WriteLine(str Like pattern)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_like_operator_empty_string_matching() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("" Like "") & "|" & ("" Like "*") & "|" & ("" Like "?"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_like_operator_null_operand_returns_false() {
    let src = r#"
Module Program
    Sub Main()
        Dim str As String = Nothing
        Console.WriteLine(str Like "A*")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_like_operator_numeric_coercion() {
    let src = r#"
Module Program
    Sub Main()
        Dim num As Integer = 12345
        Console.WriteLine(num Like "12*")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_like_operator_date_coercion() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1)
        Console.WriteLine(dt.ToString("yyyy-MM-dd") Like "2025-*")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_like_operator_hyphen_in_char_list() {
    let src = r#"
Module Program
    Sub Main()
        ' To match literal hyphen inside brackets, place it first or last!
        Console.WriteLine(("-" Like "[-abc]") & "|" & ("a" Like "[-abc]"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_like_operator_digit_range_brackets() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("5" Like "[0-9]") & "|" & ("A" Like "[0-9]"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_like_operator_multiple_bracket_groups() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("A1" Like "[A-Z][0-9]") & "|" & ("1A" Like "[A-Z][0-9]"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_like_operator_string_variable_pattern() {
    let src = r#"
Module Program
    Sub Main()
        Dim input = "file_2025.log"
        Dim pat = "file_####.log"
        Console.WriteLine(input Like pat)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_like_operator_whitespace_matching() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(("A B" Like "A?B") & "|" & ("A B" Like "A*B"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_like_operator_unclosed_bracket_matches_literal() {
    let src = r#"
Module Program
    Sub Main()
        ' Unclosed bracket in Like pattern throws or matches literally depending on runtime
        Try
            Dim res = "A[" Like "A["
            Console.WriteLine(res)
        Catch ex As System.Exception
            Console.WriteLine("Like Pattern Syntax Exception Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
