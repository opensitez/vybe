use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Try...Catch Rethrow (Throw), Custom Catch Filters
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_try_catch_rethrow_bare_throw() {
    let src = r#"
Imports System

Module Program
    Private Sub Helper()
        Try
            Dim zero As Integer = 0
            Dim res As Integer = 10 \ zero
        Catch ex As Exception
            Console.WriteLine("Logging in Helper")
            Throw ' Bare rethrow preserves stack trace
        End Try
    End Sub

    Sub Main()
        Try
            Helper()
        Catch ex As DivideByZeroException
            Console.WriteLine("Caught in Main: " & ex.GetType().Name)
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Logging in Helper", "Caught in Main: DivideByZeroException"]
    );
}

#[test]
fn test_vb_try_catch_rethrow_wrapped_exception() {
    let src = r#"
Imports System

Module Program
    Private Sub Process()
        Try
            Throw New FormatException("Invalid Number")
        Catch ex As FormatException
            Throw New InvalidOperationException("Process Failed", ex)
        End Try
    End Sub

    Sub Main()
        Try
            Process()
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message & " | Inner: " & ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Process Failed | Inner: Invalid Number"]);
}

#[test]
fn test_vb_try_catch_multiple_catch_blocks_order() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim arr As Integer() = {1, 2}
            Console.WriteLine(arr(5))
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Caught IndexOutOfRangeException")
        Catch ex As Exception
            Console.WriteLine("Caught General Exception")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Caught IndexOutOfRangeException"]);
}

#[test]
fn test_vb_try_catch_when_filter_clause() {
    let src = r#"
Imports System

Module Program
    Private Sub PerformAction(code As Integer)
        Try
            Throw New ArgumentException("Error", "ParamName")
        Catch ex As ArgumentException When code = 1
            Console.WriteLine("Handled Code 1")
        Catch ex As ArgumentException When code = 2
            Console.WriteLine("Handled Code 2")
        Catch ex As Exception
            Console.WriteLine("Handled Fallback")
        End Try
    End Sub

    Sub Main()
        PerformAction(1)
        PerformAction(2)
        PerformAction(3)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Handled Code 1", "Handled Code 2", "Handled Fallback"]
    );
}

#[test]
fn test_vb_try_finally_always_executes_on_uncaught() {
    let src = r#"
Imports System

Module Program
    Private Sub Execute()
        Try
            Console.WriteLine("Inside Try")
            Throw New InvalidOperationException()
        Finally
            Console.WriteLine("Inside Finally")
        End Try
    End Sub

    Sub Main()
        Try
            Execute()
        Catch ex As Exception
            Console.WriteLine("Caught in Main")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Inside Try", "Inside Finally", "Caught in Main"]
    );
}

#[test]
fn test_vb_try_finally_always_executes_on_normal_return() {
    let src = r#"
Module Program
    Private Function Compute() As String
        Try
            Return "Result"
        Finally
            Console.WriteLine("Cleanup in Finally")
        End Try
    End Function

    Sub Main()
        Dim res = Compute()
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Cleanup in Finally", "Result"]);
}

#[test]
fn test_vb_custom_exception_derived_class() {
    let src = r#"
Imports System

Class BusinessRuleException
    Inherits Exception
    Public Property ErrorCode As Integer
    Public Sub New(code As Integer, msg As String)
        MyBase.New(msg)
        ErrorCode = code
    End Sub
End Class

Module Program
    Sub Main()
        Try
            Throw New BusinessRuleException(404, "Entity Not Found")
        Catch ex As BusinessRuleException
            Console.WriteLine("Code: " & ex.ErrorCode & " | Msg: " & ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Code: 404 | Msg: Entity Not Found"]);
}

#[test]
fn test_vb_nested_try_catch_blocks() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Console.WriteLine("Outer Try Start")
            Try
                Throw New OverflowException("Inner Exception")
            Catch ex As OverflowException
                Console.WriteLine("Inner Catch: " & ex.Message)
            End Try
            Console.WriteLine("Outer Try End")
        Catch ex As Exception
            Console.WriteLine("Outer Catch")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec![
            "Outer Try Start",
            "Inner Catch: Inner Exception",
            "Outer Try End"
        ]
    );
}

#[test]
fn test_vb_try_catch_filter_side_effect_function() {
    let src = r#"
Imports System

Module Program
    Private Function LogAndCheck(ex As Exception) As Boolean
        Console.WriteLine("Filter Evaluated for: " & ex.Message)
        Return True
    End Function

    Sub Main()
        Try
            Throw New InvalidOperationException("TestMsg")
        Catch ex As Exception When LogAndCheck(ex)
            Console.WriteLine("Catch Executed")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Filter Evaluated for: TestMsg", "Catch Executed"]
    );
}

#[test]
fn test_vb_exception_data_dictionary() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim ex As New Exception("Data Error")
            ex.Data("TimeStamp") = "2025-01-01"
            Throw ex
        Catch ex As Exception
            Console.WriteLine("TimeStamp: " & ex.Data("TimeStamp").ToString())
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TimeStamp: 2025-01-01"]);
}

#[test]
fn test_vb_aggregate_exception_flattening() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Try
            Dim inner1 As New InvalidOperationException("Op1")
            Dim inner2 As New ArgumentException("Op2")
            Throw New AggregateException("Batch Failed", inner1, inner2)
        Catch ex As AggregateException
            For Each inner In ex.InnerExceptions
                Console.WriteLine("Inner: " & inner.Message)
            Next
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Inner: Op1", "Inner: Op2"]);
}

#[test]
fn test_vb_throw_null_reference_throws_null_pointer() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim nullEx As Exception = Nothing
            Throw nullEx
        Catch ex As NullReferenceException
            Console.WriteLine("Caught NullReferenceException on throw null")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Caught NullReferenceException on throw null"]
    );
}

#[test]
fn test_vb_catch_assigns_to_variable() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim capturedEx As Exception = Nothing
        Try
            Throw New Exception("Captured")
        Catch ex As Exception
            capturedEx = ex
        End Try
        Console.WriteLine(capturedEx.Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Captured"]);
}

#[test]
fn test_vb_try_catch_in_constructor() {
    let src = r#"
Imports System

Class ConstructFailure
    Public Sub New()
        Try
            Throw New InvalidOperationException("Fail in New")
        Catch ex As Exception
            Console.WriteLine("Handled in New")
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New ConstructFailure()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Handled in New"]);
}

#[test]
fn test_vb_try_catch_in_property_getter() {
    let src = r#"
Imports System

Class SafeProperty
    Public ReadOnly Property SafeValue As Integer
        Get
            Try
                Dim zero As Integer = 0
                Return 100 \ zero
            Catch ex As DivideByZeroException
                Return -1
            End Try
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim sp As New SafeProperty()
        Console.WriteLine(sp.SafeValue)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1"]);
}

#[test]
fn test_vb_try_catch_in_shared_constructor() {
    let src = r#"
Imports System

Class SafeSharedInit
    Public Shared Loaded As Boolean = False
    Shared Sub New()
        Try
            Throw New Exception("Init Fail")
        Catch ex As Exception
            Console.WriteLine("Caught in Shared Sub New")
            Loaded = True
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(SafeSharedInit.Loaded)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Caught in Shared Sub New", "True"]);
}

#[test]
fn test_vb_try_finally_rethrow_unhandled() {
    let src = r#"
Imports System

Module Program
    Private Sub Outer()
        Try
            Inner()
        Catch ex As Exception
            Console.WriteLine("Outer Caught: " & ex.Message)
        End Try
    End Sub

    Private Sub Inner()
        Try
            Throw New Exception("Deep Error")
        Finally
            Console.WriteLine("Inner Finally Executed")
        End Try
    End Sub

    Sub Main()
        Outer()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Inner Finally Executed", "Outer Caught: Deep Error"]
    );
}

#[test]
fn test_vb_try_catch_expression_bodied_function_call() {
    let src = r#"
Imports System

Module Program
    Private Function ParseSafely(input As String) As Integer
        Try
            Return Integer.Parse(input)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Sub Main()
        Console.WriteLine(ParseSafely("123") & "|" & ParseSafely("abc"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123|0"]);
}

#[test]
fn test_vb_try_catch_exit_sub_finally_still_runs() {
    let src = r#"
Module Program
    Private Sub TestExit()
        Try
            Console.WriteLine("Step 1")
            Exit Sub
        Finally
            Console.WriteLine("Finally on Exit Sub")
        End Try
    End Sub

    Sub Main()
        TestExit()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Step 1", "Finally on Exit Sub"]);
}

#[test]
fn test_vb_try_catch_exit_try_statement() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Console.WriteLine("Before Exit Try")
            Exit Try
            Console.WriteLine("After Exit Try")
        Catch ex As Exception
            Console.WriteLine("Catch Block")
        Finally
            Console.WriteLine("Finally Block")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Before Exit Try", "Finally Block"]);
}
