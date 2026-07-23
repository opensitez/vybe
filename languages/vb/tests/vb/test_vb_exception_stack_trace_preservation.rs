use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Exception Stack Trace Preservation & Inner Exception
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_exception_stack_trace_populated() {
    let src = r#"
Imports System

Module Program
    Private Sub ThrowError()
        Throw New InvalidOperationException("Stack Trace Test")
    End Sub

    Sub Main()
        Try
            ThrowError()
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace IsNot Nothing)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_exception_source_property() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim ex As New Exception("Custom Source")
            ex.Source = "MyCustomModule"
            Throw ex
        Catch ex As Exception
            Console.WriteLine(ex.Source)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["MyCustomModule"]);
}

#[test]
fn test_vb_exception_hresult_property() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Throw New FormatException()
        Catch ex As Exception
            Console.WriteLine(ex.HResult <> 0)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_exception_target_site_method_reflection() {
    let src = r#"
Imports System

Module Program
    Private Sub FailingMethod()
        Throw New InvalidOperationException("TargetSite")
    End Sub

    Sub Main()
        Try
            FailingMethod()
        Catch ex As Exception
            Console.WriteLine(ex.TargetSite IsNot Nothing AndAlso ex.TargetSite.Name = "FailingMethod")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_inner_exception_chain_three_levels() {
    let src = r#"
Imports System

Module Program
    Private Sub Level1()
        Throw New ArgumentException("Level 1 Error")
    End Sub

    Private Sub Level2()
        Try
            Level1()
        Catch ex As Exception
            Throw New InvalidOperationException("Level 2 Wrapper", ex)
        End Try
    End Sub

    Sub Main()
        Try
            Level2()
        Catch ex As Exception
            Console.WriteLine(ex.Message & " -> " & ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Level 2 Wrapper -> Level 1 Error"]);
}

#[test]
fn test_vb_exception_get_base_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim root As New OverflowException("Root Cause")
            Dim mid As New InvalidOperationException("Mid Layer", root)
            Dim top As New Exception("Top Layer", mid)
            Throw top
        Catch ex As Exception
            Console.WriteLine(ex.GetBaseException().Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Root Cause"]);
}

#[test]
fn test_vb_exception_to_string_formatting() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Throw New InvalidOperationException("Formatting Test")
        Catch ex As Exception
            Dim str = ex.ToString()
            Console.WriteLine(str.Contains("InvalidOperationException") AndAlso str.Contains("Formatting Test"))
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_exception_dispatch_info_capture_throw() {
    let src = r#"
Imports System
Imports System.Runtime.ExceptionServices

Module Program
    Private captured As ExceptionDispatchInfo

    Private Sub CauseError()
        Try
            Throw New InvalidOperationException("Captured Exception")
        Catch ex As Exception
            captured = ExceptionDispatchInfo.Capture(ex)
        End Try
    End Sub

    Sub Main()
        CauseError()
        Try
            captured.Throw()
        Catch ex As Exception
            Console.WriteLine("Re-thrown: " & ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Re-thrown: Captured Exception"]);
}

#[test]
fn test_vb_custom_exception_constructor_overloads() {
    let src = r#"
Imports System

Class CustomException
    Inherits Exception
    Public Sub New() : MyBase.New("Default Message") : End Sub
    Public Sub New(msg As String) : MyBase.New(msg) : End Sub
    Public Sub New(msg As String, inner As Exception) : MyBase.New(msg, inner) : End Sub
End Class

Module Program
    Sub Main()
        Dim e1 As New CustomException()
        Dim e2 As New CustomException("Custom Msg")
        Dim e3 As New CustomException("Wrapper", e1)
        Console.WriteLine(e1.Message & "|" & e2.Message & "|" & e3.InnerException.Message)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Default Message|Custom Msg|Default Message"]
    );
}

#[test]
fn test_vb_exception_help_link_property() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim ex As New Exception("Check Docs")
            ex.HelpLink = "https://docs.microsoft.com"
            Throw ex
        Catch ex As Exception
            Console.WriteLine(ex.HelpLink)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["https://docs.microsoft.com"]);
}

#[test]
fn test_vb_exception_in_async_task_run() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Try
            Dim t = Task.Run(Sub() Throw New InvalidOperationException("Async Error"))
            t.Wait()
        Catch ex As AggregateException
            Console.WriteLine("Async Caught: " & ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Async Caught: Async Error"]);
}

#[test]
fn test_vb_exception_type_casting_hierarchy() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Throw New ArgumentNullException("paramName", "Null Param")
        Catch ex As ArgumentException
            Console.WriteLine("Caught as ArgumentException: " & ex.Message.Split(vbCrLf(0))(0))
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Caught as ArgumentException: Null Param"]);
}

#[test]
fn test_vb_exception_in_iterator_function_yield() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Private Iterator Function GenerateItems() As IEnumerable(Of Integer)
        Yield 1
        Yield 2
        Throw New InvalidOperationException("Generator Error")
    End Function

    Sub Main()
        Try
            For Each item In GenerateItems()
                Console.WriteLine("Item: " & item)
            Next
        Catch ex As Exception
            Console.WriteLine("Caught in Iterator Loop: " & ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec![
            "Item: 1",
            "Item: 2",
            "Caught in Iterator Loop: Generator Error"
        ]
    );
}

#[test]
fn test_vb_exception_in_delegate_multicast_stops_chain() {
    let src = r#"
Imports System

Module Program
    Private Sub First()
        Console.WriteLine("First Executed")
        Throw New InvalidOperationException("First Failed")
    End Sub
    Private Sub Second()
        Console.WriteLine("Second Executed")
    End Sub

    Sub Main()
        Dim act As Action = AddressOf First
        act = CType([Delegate].Combine(act, New Action(AddressOf Second)), Action)
        Try
            act()
        Catch ex As Exception
            Console.WriteLine("Caught: " & ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First Executed", "Caught: First Failed"]);
}

#[test]
fn test_vb_exception_in_event_handler_chain() {
    let src = r#"
Imports System

Class Publisher
    Public Event Fire As Action
    Public Sub RaiseFire()
        RaiseEvent Fire()
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Publisher()
        AddHandler p.Fire, Sub() Console.WriteLine("Handler 1")
        AddHandler p.Fire, Sub() Throw New Exception("Handler 2 Failed")
        AddHandler p.Fire, Sub() Console.WriteLine("Handler 3")

        Try
            p.RaiseFire()
        Catch ex As Exception
            Console.WriteLine("Caught in Main: " & ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Handler 1", "Caught in Main: Handler 2 Failed"]
    );
}

#[test]
fn test_vb_exception_stack_trace_line_number_info() {
    let src = r#"
Imports System

Module Program
    Private Sub DeepFail()
        Throw New Exception("Deep Error")
    End Sub

    Sub Main()
        Try
            DeepFail()
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace.Contains("DeepFail"))
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_exception_thrown_from_lambda_closure() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim lambda As Action = Sub() Throw New InvalidOperationException("Closure Error")
        Try
            lambda()
        Catch ex As Exception
            Console.WriteLine("Caught Lambda: " & ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Caught Lambda: Closure Error"]);
}

#[test]
fn test_vb_exception_rethrow_preserves_original_stack_frames() {
    let src = r#"
Imports System

Module Program
    Private Sub Level1()
        Throw New InvalidOperationException("Root")
    End Sub

    Private Sub Level2()
        Try
            Level1()
        Catch ex As Exception
            Throw ' Preserves Level1 in stack trace
        End Try
    End Sub

    Sub Main()
        Try
            Level2()
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace.Contains("Level1"))
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_exception_throw_ex_resets_stack_trace() {
    let src = r#"
Imports System

Module Program
    Private Sub Level1()
        Throw New InvalidOperationException("Root")
    End Sub

    Private Sub Level2()
        Try
            Level1()
        Catch ex As Exception
            Throw ex ' Re-throwing ex resets the stack trace to Level2
        End Try
    End Sub

    Sub Main()
        Try
            Level2()
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace.Contains("Level2"))
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_exception_in_destructor_finalizer_suppressed() {
    let src = r#"
Imports System

Class Destructible
    Protected Overrides Sub Finalize()
        Try
            ' Destructors should never throw uncaught exceptions
            Console.WriteLine("Finalizer executed")
        Catch ex As Exception
        Finally
            MyBase.Finalize()
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New Destructible()
        d = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Finalizer executed"]);
}
