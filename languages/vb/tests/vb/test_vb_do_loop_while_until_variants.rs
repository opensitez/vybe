use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Do...Loop Variants (While/Until at Top/Bottom)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_do_while_top_condition() {
    let src = r#"
Module Program
    Sub Main()
        Dim count = 0
        Do While count < 3
            count += 1
        Loop
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_do_until_top_condition() {
    let src = r#"
Module Program
    Sub Main()
        Dim count = 0
        Do Until count >= 3
            count += 1
        Loop
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_do_loop_while_bottom_condition() {
    let src = r#"
Module Program
    Sub Main()
        Dim count = 0
        Do
            count += 1
        Loop While count < 3
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_do_loop_until_bottom_condition() {
    let src = r#"
Module Program
    Sub Main()
        Dim count = 0
        Do
            count += 1
        Loop Until count >= 3
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_do_loop_bottom_condition_executes_at_least_once() {
    let src = r#"
Module Program
    Sub Main()
        Dim count = 10
        Do
            count += 1
        Loop While count < 5 ' False on first check, but loop body ran once!
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["11"]);
}

#[test]
fn test_vb_do_while_top_false_never_executes() {
    let src = r#"
Module Program
    Sub Main()
        Dim count = 10
        Do While count < 5 ' False initially
            count += 1
        Loop
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_do_loop_exit_do() {
    let src = r#"
Module Program
    Sub Main()
        Dim count = 0
        Do
            count += 1
            If count = 3 Then Exit Do
        Loop While count < 10
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_nested_do_loops_exit_inner() {
    let src = r#"
Module Program
    Sub Main()
        Dim outerCount = 0
        Dim innerCount = 0
        Do While outerCount < 2
            outerCount += 1
            Do
                innerCount += 1
                If innerCount >= 2 Then Exit Do
            Loop While True
        Loop
        Console.WriteLine(outerCount & "|" & innerCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|3"]);
}

#[test]
fn test_vb_do_loop_with_continue_do() {
    let src = r#"
Module Program
    Sub Main()
        Dim sum = 0
        Dim i = 0
        Do While i < 5
            i += 1
            If i = 3 Then Continue Do
            sum += i
        Loop
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12"]);
}

#[test]
fn test_vb_do_loop_infinite_with_exit_do_break() {
    let src = r#"
Module Program
    Sub Main()
        Dim i = 0
        Do
            i += 5
            If i > 20 Then Exit Do
        Loop
        Console.WriteLine(i)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["25"]);
}

#[test]
fn test_vb_do_while_complex_boolean_expression() {
    let src = r#"
Module Program
    Sub Main()
        Dim a = 0
        Dim b = 10
        Do While a < 5 AndAlso b > 5
            a += 1
            b -= 1
        Loop
        Console.WriteLine(a & "|" & b)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5|5"]);
}

#[test]
fn test_vb_do_until_complex_boolean_expression() {
    let src = r#"
Module Program
    Sub Main()
        Dim x = 0
        Do Until x >= 10 OrElse x = 5
            x += 1
        Loop
        Console.WriteLine(x)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_while_wend_legacy_loop() {
    let src = r#"
Module Program
    Sub Main()
        Dim i = 0
        While i < 3
            i += 1
        End While
        Console.WriteLine(i)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_do_loop_with_byref_mutation_inside() {
    let src = r#"
Module Program
    Private Sub Increment(ByRef val As Integer)
        val += 1
    End Sub

    Sub Main()
        Dim count = 0
        Do
            Increment(count)
        Loop Until count = 4
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_do_loop_reading_array_elements() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30, 40}
        Dim idx = 0
        Dim sum = 0
        Do While idx < numbers.Length
            sum += numbers(idx)
            idx += 1
        Loop
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_do_loop_with_try_catch_inside() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim attempts = 0
        Dim successCount = 0
        Do While attempts < 3
            attempts += 1
            Try
                If attempts = 2 Then Throw New Exception("Transient Error")
                successCount += 1
            Catch ex As Exception
                Console.WriteLine("Error Handled at Attempt " & attempts)
            End Try
        Loop
        Console.WriteLine(successCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Error Handled at Attempt 2", "2"]);
}

#[test]
fn test_vb_do_loop_modifying_step_variable() {
    let src = r#"
Module Program
    Sub Main()
        Dim val = 1
        Do While val < 100
            val *= 2
        Loop
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["128"]);
}

#[test]
fn test_vb_do_until_bottom_condition_executes_body_once() {
    let src = r#"
Module Program
    Sub Main()
        Dim count = 100
        Do
            count += 5
        Loop Until count > 50 ' True after first check, loop ends!
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["105"]);
}

#[test]
fn test_vb_do_loop_with_return_statement_inside() {
    let src = r#"
Module Program
    Private Function FindTarget() As Integer
        Dim i = 0
        Do
            i += 1
            If i = 5 Then Return i * 10
        Loop While True
        Return -1
    End Function

    Sub Main()
        Console.WriteLine(FindTarget())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_do_loop_flag_controlled() {
    let src = r#"
Module Program
    Sub Main()
        Dim running = True
        Dim stepCount = 0
        Do While running
            stepCount += 1
            If stepCount = 4 Then running = False
        Loop
        Console.WriteLine(stepCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}
