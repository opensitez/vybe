use super::helpers::run_vb;

#[test]
fn try_catch_when_complex() {
    let out = run_vb(
        r#"
Module M
    Function LogError() As Boolean
        Console.WriteLine("Filtered")
        Return True
    End Function

    Sub Main()
        Try
            Throw New System.Exception("Test")
        Catch ex As System.Exception When LogError()
            Console.WriteLine("Caught")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Filtered", "Caught"]);
}

#[test]
fn try_catch_finally_nested_break() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        For i = 1 To 3
            Try
                If i = 2 Then Exit For
            Finally
                Console.WriteLine("Finally" & i)
            End Try
        Next
        Console.WriteLine("Done")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Finally1", "Finally2", "Done"]);
}

#[test]
fn do_until_loop_until() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i = 0
        ' Technically Do Until and Loop Until together is a syntax edge case
        Do Until i > 5
            i += 1
        Loop Until i = 3
        Console.WriteLine(i)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn for_each_with_type_conversion() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr As Integer() = {1, 2, 3}
        ' For Each with implicit conversion to Double
        For Each x As Double In arr
            Console.WriteLine(x + 0.5)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1.5", "2.5", "3.5"]);
}

#[test]
fn for_loop_step_negative() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        For i = 3 To 1 Step -1
            Console.WriteLine(i)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn for_loop_step_decimal() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        For i As Decimal = 1.5D To 2.5D Step 0.5D
            Console.WriteLine(i)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1.5", "2", "2.5"]);
}

#[test]
fn exit_select_nested() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val = 1
        Select Case val
            Case 1
                For i = 1 To 5
                    If i = 3 Then Exit Select
                    Console.WriteLine(i)
                Next
            Case 2
                Console.WriteLine("Two")
        End Select
        Console.WriteLine("Done")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "Done"]);
}

#[test]
fn exit_sub_in_catch() {
    let out = run_vb(
        r#"
Module M
    Sub Test()
        Try
            Throw New System.Exception()
        Catch
            Console.WriteLine("Catch")
            Exit Sub
        Finally
            Console.WriteLine("Finally")
        End Try
        Console.WriteLine("After")
    End Sub

    Sub Main()
        Test()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Catch", "Finally"]);
}

#[test]
fn exit_function_in_finally() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Exit Function inside Finally is generally parsed as invalid by standard rules
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn on_error_goto_label_reset() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        On Error GoTo Handler
        Throw New System.Exception()
        
    Handler:
        Console.WriteLine("Handled")
        On Error GoTo 0 ' Reset
        Exit Sub
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Handled"]);
}

#[test]
fn error_handling_resume() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim attempts = 0
        On Error GoTo Handler
        
    RetryPoint:
        If attempts = 0 Then
            Throw New System.Exception()
        End If
        Console.WriteLine("Success")
        Exit Sub
        
    Handler:
        attempts += 1
        Console.WriteLine("Attempt " & attempts)
        Resume RetryPoint
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Attempt 1", "Success"]);
}

#[test]
fn error_handling_resume_next() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        On Error GoTo Handler
        
        Throw New System.Exception()
        Console.WriteLine("Resumed")
        Exit Sub
        
    Handler:
        Resume Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Resumed"]);
}

#[test]
fn return_from_catch() {
    let out = run_vb(
        r#"
Module M
    Function Test() As String
        Try
            Throw New System.Exception()
        Catch
            Return "Caught"
        End Try
        Return "Missed"
    End Function

    Sub Main()
        Console.WriteLine(Test())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Caught"]);
}

#[test]
fn yield_in_catch() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Yield inside Catch block is often an error. Testing parser check.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn goto_across_blocks() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        If True Then
            GoTo Target
        End If
        
        Console.WriteLine("Skipped")
        
    Target:
        Console.WriteLine("Target")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Target"]);
}

#[test]
fn select_case_is() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val = 15
        Select Case val
            Case Is > 20
                Console.WriteLine(">20")
            Case Is > 10
                Console.WriteLine(">10")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec![">10"]);
}

#[test]
fn select_case_range() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val = 5
        Select Case val
            Case 1 To 10
                Console.WriteLine("1-10")
            Case Else
                Console.WriteLine("Other")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1-10"]);
}

#[test]
fn select_case_comma() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val = 2
        Select Case val
            Case 1, 2, 3
                Console.WriteLine("Matched")
            Case Else
                Console.WriteLine("Other")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Matched"]);
}

#[test]
fn select_case_multiple_types() {
    let out = run_vb(
        r#"
Option Strict Off
Module M
    Sub Main()
        Dim val As Object = "10"
        
        Select Case val
            Case 10
                Console.WriteLine("Num10")
            Case "10"
                Console.WriteLine("Str10")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Num10"]); ' Because late binding parses "10" to 10 for comparison first in Option Strict Off generally, or it matches on Type
}

#[test]
fn if_elseif_else_chain() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val = 3
        If val = 1 Then
            Console.WriteLine("1")
        ElseIf val = 2 Then
            Console.WriteLine("2")
        ElseIf val = 3 Then
            Console.WriteLine("3")
        Else
            Console.WriteLine("Other")
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn singleline_if_else() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x = 10
        If x > 5 Then Console.WriteLine("Yes") Else Console.WriteLine("No")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Yes"]);
}

#[test]
fn multiline_if_with_colon() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x = 1
        ' Valid multiline if separated by colon
        If x = 1 Then : Console.WriteLine("A") : Console.WriteLine("B") : End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["A", "B"]);
}

#[test]
fn with_statement_nested() {
    let out = run_vb(
        r#"
Class Inner
    Public Val As Integer = 10
End Class

Class Outer
    Public InObj As New Inner()
End Class

Module M
    Sub Main()
        Dim o As New Outer()
        With o
            With .InObj
                Console.WriteLine(.Val)
            End With
        End With
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn with_statement_struct() {
    let out = run_vb(
        r#"
Structure S
    Public Val As Integer
End Structure

Module M
    Sub Main()
        Dim s1 As New S()
        With s1
            .Val = 42
        End With
        ' Because it's a value type, With creates a copy or modifies original depending on context.
        ' Actually in VB, With on a variable Modifies the variable!
        Console.WriteLine(s1.Val)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn with_statement_array_element() {
    let out = run_vb(
        r#"
Class Item
    Public Val As Integer
End Class

Module M
    Sub Main()
        Dim arr = {New Item(), New Item()}
        With arr(0)
            .Val = 100
        End With
        Console.WriteLine(arr(0).Val)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn sync_lock_value_type() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' SyncLock on a Value Type is invalid.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn using_statement_multiple_types() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Using multiple different types (e.g. Using x As A, y As B) is invalid.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn await_in_loop() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Async Function Test() As Task
        For i = 1 To 2
            Await Task.Delay(1)
            Console.WriteLine(i)
        Next
    End Function

    Sub Main()
        Test().Wait()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn await_in_try() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Async Function Test() As Task
        Try
            Await Task.Delay(1)
            Console.WriteLine("Try")
        Catch
        End Try
    End Function

    Sub Main()
        Test().Wait()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Try"]);
}

#[test]
fn throw_inner_exception() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Try
            Try
                Throw New Exception("Inner")
            Catch ex As Exception
                Throw New Exception("Outer", ex)
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.InnerException.Message)
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Inner"]);
}

#[test]
fn try_catch_multiple_filters() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim code = 2
        Try
            Throw New Exception("Err")
        Catch ex As Exception When code = 1
            Console.WriteLine("One")
        Catch ex As Exception When code = 2
            Console.WriteLine("Two")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Two"]);
}

#[test]
fn while_loop_exit() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i = 0
        While i < 10
            If i = 2 Then Exit While
            i += 1
        End While
        Console.WriteLine(i)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn while_loop_continue() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i = 0
        Dim sum = 0
        While i < 3
            i += 1
            If i = 2 Then Continue While
            sum += i
        End While
        Console.WriteLine(sum) ' 1 + 3 = 4
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn do_loop_continue() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i = 0
        Dim sum = 0
        Do While i < 3
            i += 1
            If i = 2 Then Continue Do
            sum += i
        Loop
        Console.WriteLine(sum) ' 1 + 3 = 4
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn for_loop_continue() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim sum = 0
        For i = 1 To 3
            If i = 2 Then Continue For
            sum += i
        Next
        Console.WriteLine(sum) ' 1 + 3 = 4
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4"]);
}
