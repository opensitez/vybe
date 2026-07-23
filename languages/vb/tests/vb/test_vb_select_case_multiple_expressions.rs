use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Select Case Expressions, Ranges (To), Is Clauses & Lists
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_select_case_comma_separated_values() {
    let src = r#"
Module Program
    Sub Main()
        Dim day = 6
        Select Case day
            Case 1, 7
                Console.WriteLine("Weekend")
            Case 2, 3, 4, 5, 6
                Console.WriteLine("Weekday")
            Case Else
                Console.WriteLine("Invalid")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Weekday"]);
}

#[test]
fn test_vb_select_case_range_to_clause() {
    let src = r#"
Module Program
    Sub Main()
        Dim score = 85
        Select Case score
            Case 90 To 100
                Console.WriteLine("A")
            Case 80 To 89
                Console.WriteLine("B")
            Case 70 To 79
                Console.WriteLine("C")
            Case Else
                Console.WriteLine("F")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B"]);
}

#[test]
fn test_vb_select_case_is_relational_operator() {
    let src = r#"
Module Program
    Sub Main()
        Dim val = 150
        Select Case val
            Case Is < 0
                Console.WriteLine("Negative")
            Case Is >= 100
                Console.WriteLine("High")
            Case Else
                Console.WriteLine("Normal")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["High"]);
}

#[test]
fn test_vb_select_case_string_pattern_matching() {
    let src = r#"
Module Program
    Sub Main()
        Dim fruit = "Banana"
        Select Case fruit
            Case "Apple", "Pear"
                Console.WriteLine("Pome Fruit")
            Case "Banana", "Mango", "Pineapple"
                Console.WriteLine("Tropical Fruit")
            Case Else
                Console.WriteLine("Other")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Tropical Fruit"]);
}

#[test]
fn test_vb_select_case_type_checking_with_typeof() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj As Object = "Hello"
        Select Case True
            Case TypeOf obj Is String
                Console.WriteLine("IsString")
            Case TypeOf obj Is Integer
                Console.WriteLine("IsInteger")
            Case Else
                Console.WriteLine("OtherType")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IsString"]);
}

#[test]
fn test_vb_select_case_enum_values() {
    let src = r#"
Enum LogLevel
    Debug = 1
    Info = 2
    Error = 3
End Enum

Module Program
    Sub Main()
        Dim level = LogLevel.Error
        Select Case level
            Case LogLevel.Debug, LogLevel.Info
                Console.WriteLine("Non-Critical")
            Case LogLevel.Error
                Console.WriteLine("Critical")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Critical"]);
}

#[test]
fn test_vb_select_case_char_ranges() {
    let src = r#"
Module Program
    Sub Main()
        Dim ch As Char = "k"c
        Select Case ch
            Case "a"c To "m"c
                Console.WriteLine("First Half")
            Case "n"c To "z"c
                Console.WriteLine("Second Half")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First Half"]);
}

#[test]
fn test_vb_select_case_mixed_clause_expressions() {
    let src = r#"
Module Program
    Sub Main()
        Dim x = 5
        Select Case x
            Case 1, 3, 7 To 10, Is > 100
                Console.WriteLine("Matched Group A")
            Case 4 To 6, Is < 0
                Console.WriteLine("Matched Group B")
            Case Else
                Console.WriteLine("Default")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Matched Group B"]);
}

#[test]
fn test_vb_select_case_first_matching_case_executes_only() {
    let src = r#"
Module Program
    Sub Main()
        Dim val = 10
        Select Case val
            Case 10
                Console.WriteLine("First 10")
            Case Is >= 10
                Console.WriteLine("Second >= 10")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First 10"]);
}

#[test]
fn test_vb_select_case_evaluates_select_expression_once() {
    let src = r#"
Module Program
    Private Function GetValue(ByRef count As Integer) As Integer
        count += 1
        Return 5
    End Function

    Sub Main()
        Dim evalCount = 0
        Select Case GetValue(evalCount)
            Case 1
                Console.WriteLine("One")
            Case 5
                Console.WriteLine("Five|Evals=" & evalCount)
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Five|Evals=1"]);
}

#[test]
fn test_vb_select_case_case_else_fallback() {
    let src = r#"
Module Program
    Sub Main()
        Dim color = "Purple"
        Select Case color
            Case "Red"
                Console.WriteLine("R")
            Case "Blue"
                Console.WriteLine("B")
            Case Else
                Console.WriteLine("Fallback")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fallback"]);
}

#[test]
fn test_vb_select_case_option_compare_text_matching() {
    let src = r#"
Option Compare Text

Module Program
    Sub Main()
        Dim cmd = "EXIT"
        Select Case cmd
            Case "exit", "quit"
                Console.WriteLine("Stopping")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Stopping"]);
}

#[test]
fn test_vb_select_case_boolean_true_pattern_matching() {
    let src = r#"
Module Program
    Sub Main()
        Dim age = 25
        Dim isStudent = True
        Select Case True
            Case age < 18
                Console.WriteLine("Minor")
            Case age >= 18 AndAlso isStudent
                Console.WriteLine("Student Adult")
            Case Else
                Console.WriteLine("Adult")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Student Adult"]);
}

#[test]
fn test_vb_select_case_decimal_values() {
    let src = r#"
Module Program
    Sub Main()
        Dim price As Decimal = 19.99D
        Select Case price
            Case 0.0D To 9.99D
                Console.WriteLine("Low")
            Case 10.0D To 49.99D
                Console.WriteLine("Medium")
            Case Else
                Console.WriteLine("High")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Medium"]);
}

#[test]
fn test_vb_select_case_nested_select_statements() {
    let src = r#"
Module Program
    Sub Main()
        Dim category = "Tech"
        Dim subCat = "Mobile"
        Select Case category
            Case "Tech"
                Select Case subCat
                    Case "Mobile"
                        Console.WriteLine("Tech-Mobile")
                    Case "Desktop"
                        Console.WriteLine("Tech-Desktop")
                End Select
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Tech-Mobile"]);
}

#[test]
fn test_vb_select_case_expression_function_call_in_case() {
    let src = r#"
Module Program
    Private Function DoubleVal(x As Integer) As Integer
        Return x * 2
    End Function

    Sub Main()
        Dim val = 10
        Select Case val
            Case DoubleVal(5)
                Console.WriteLine("Matched DoubleVal(5)")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Matched DoubleVal(5)"]);
}

#[test]
fn test_vb_select_case_nullable_value_type() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim n As Integer? = 42
        Select Case n
            Case HasValue
                Console.WriteLine("Value: " & n.Value)
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Value: 42"]);
}

#[test]
fn test_vb_select_case_with_exit_select() {
    let src = r#"
Module Program
    Sub Main()
        Dim x = 1
        Select Case x
            Case 1
                Console.WriteLine("Start Case 1")
                If x = 1 Then Exit Select
                Console.WriteLine("End Case 1")
        End Select
        Console.WriteLine("After Select")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Start Case 1", "After Select"]);
}

#[test]
fn test_vb_select_case_is_not_equal_operator() {
    let src = r#"
Module Program
    Sub Main()
        Dim status = 200
        Select Case status
            Case Is <> 200
                Console.WriteLine("Error Status")
            Case Else
                Console.WriteLine("OK Status")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OK Status"]);
}

#[test]
fn test_vb_select_case_date_time_ranges() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 6, 15)
        Select Case dt
            Case New DateTime(2025, 1, 1) To New DateTime(2025, 12, 31)
                Console.WriteLine("Year 2025")
        End Select
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Year 2025"]);
}
