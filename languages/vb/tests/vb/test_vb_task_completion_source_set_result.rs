use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: TaskCompletionSource(Of T) Mechanics & State Transitions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_tcs_set_result_basic() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of String)()
        tcs.SetResult("ResultFromTCS")
        Console.WriteLine(tcs.Task.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ResultFromTCS"]);
}

#[test]
fn test_vb_tcs_try_set_result_first_wins() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        Dim s1 = tcs.TrySetResult(100)
        Dim s2 = tcs.TrySetResult(200)
        Console.WriteLine(s1 & "|" & s2 & "|" & tcs.Task.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|100"]);
}

#[test]
fn test_vb_tcs_set_exception_captures_failure() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        tcs.SetException(New InvalidOperationException("TCS Failed"))
        Try
            Dim x = tcs.Task.Result
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TCS Failed"]);
}

#[test]
fn test_vb_tcs_try_set_exception() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of String)()
        Dim ok = tcs.TrySetException(New Exception("First Error"))
        Dim dup = tcs.TrySetException(New Exception("Second Error"))
        Console.WriteLine(ok & "|" & dup & "|" & tcs.Task.IsFaulted)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|True"]);
}

#[test]
fn test_vb_tcs_set_canceled() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Double)()
        tcs.SetCanceled()
        Console.WriteLine(tcs.Task.IsCanceled & "|" & tcs.Task.IsCompleted)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_tcs_try_set_canceled() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of String)()
        Dim c1 = tcs.TrySetCanceled()
        Dim c2 = tcs.TrySetResult("Late")
        Console.WriteLine(c1 & "|" & c2 & "|" & tcs.Task.IsCanceled)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|True"]);
}

#[test]
fn test_vb_tcs_async_await_bridging() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function BridgeAsync(tcs As TaskCompletionSource(Of String)) As Task(Of String)
        Dim res = Await tcs.Task
        Return "Bridged: " & res
    End Function

    Sub Main()
        Dim tcs As New TaskCompletionSource(Of String)()
        Dim bgTask = BridgeAsync(tcs)
        tcs.SetResult("Payload")
        Console.WriteLine(bgTask.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bridged: Payload"]);
}

#[test]
fn test_vb_tcs_task_creation_options_run_continuations_asynchronously() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
        tcs.SetResult(True)
        Console.WriteLine(tcs.Task.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_tcs_with_custom_class_payload() {
    let src = r#"
Imports System.Threading.Tasks

Class Response
    Public Status As String
    Public Sub New(s As String) : Status = s : End Sub
End Class

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Response)()
        tcs.SetResult(New Response("200 OK"))
        Console.WriteLine(tcs.Task.Result.Status)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["200 OK"]);
}

#[test]
fn test_vb_tcs_with_value_type_struct() {
    let src = r#"
Imports System.Threading.Tasks

Structure Coordinates
    Public Lat As Double
    Public Lon As Double
    Public Sub New(l1 As Double, l2 As Double) : Lat = l1 : Lon = l2 : End Sub
End Structure

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Coordinates)()
        tcs.SetResult(New Coordinates(47.6, -122.3))
        Console.WriteLine(tcs.Task.Result.Lat & "," & tcs.Task.Result.Lon)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["47.6,-122.3"]);
}

#[test]
fn test_vb_tcs_with_tuple_payload() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of (Code As Integer, Msg As String))()
        tcs.SetResult((404, "Not Found"))
        Console.WriteLine(tcs.Task.Result.Code & " " & tcs.Task.Result.Msg)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["404 Not Found"]);
}

#[test]
fn test_vb_tcs_cancellation_token_parameter() {
    let src = r#"
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        Dim canceled = tcs.TrySetCanceled(cts.Token)
        Console.WriteLine(canceled & "|" & tcs.Task.IsCanceled)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_tcs_event_callback_resolution() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Class Producer
    Public Event DataAvailable As Action(Of String)
    Public Sub Produce(data As String)
        RaiseEvent DataAvailable(data)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Producer()
        Dim tcs As New TaskCompletionSource(Of String)()
        AddHandler p.DataAvailable, Sub(d) tcs.SetResult(d)

        p.Produce("Event Data")
        Console.WriteLine(tcs.Task.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Event Data"]);
}

#[test]
fn test_vb_tcs_status_before_completion() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        Console.WriteLine(tcs.Task.Status.ToString())
        tcs.SetResult(1)
        Console.WriteLine(tcs.Task.Status.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["WaitingForActivation", "RanToCompletion"]);
}

#[test]
fn test_vb_tcs_multiple_exceptions_enumerable() {
    let src = r#"
Imports System
Imports System.Collections.Generic
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        Dim errs As New List(Of Exception) From {
            New InvalidOperationException("E1"),
            New ArgumentException("E2")
        }
        tcs.SetException(errs)
        Console.WriteLine(tcs.Task.Exception.InnerExceptions.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_tcs_void_task_simulation() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Object)()
        tcs.SetResult(Nothing)
        Console.WriteLine(tcs.Task.IsCompleted & "|" & (tcs.Task.Result Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_tcs_continue_with_chain() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        Dim continuation = tcs.Task.ContinueWith(Function(t) t.Result * 10)
        tcs.SetResult(5)
        Console.WriteLine(continuation.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_tcs_task_timeout_race() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function GetWithTimeoutAsync(tcs As TaskCompletionSource(Of String), timeoutMs As Integer) As Task(Of String)
        Dim delayTask = Task.Delay(timeoutMs)
        Dim completed = Await Task.WhenAny(tcs.Task, delayTask)
        If completed Is tcs.Task Then
            Return Await tcs.Task
        Else
            Return "Timeout"
        End If
    End Function

    Sub Main()
        Dim tcs As New TaskCompletionSource(Of String)()
        Dim t = GetWithTimeoutAsync(tcs, 5)
        ' Don't resolve tcs, let timeout win
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Timeout"]);
}

#[test]
fn test_vb_tcs_set_result_double_invocation_throws() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        tcs.SetResult(10)
        Try
            tcs.SetResult(20)
        Catch ex As InvalidOperationException
            Console.WriteLine("Duplicate SetResult Throws InvalidOperationException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Duplicate SetResult Throws InvalidOperationException"]
    );
}

#[test]
fn test_vb_tcs_in_generic_factory_method() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Function CreateCompleted(Of T)(val As T) As Task(Of T)
        Dim tcs As New TaskCompletionSource(Of T)()
        tcs.SetResult(val)
        Return tcs.Task
    End Function

    Sub Main()
        Dim t1 = CreateCompleted(100)
        Dim t2 = CreateCompleted("Hello")
        Console.WriteLine(t1.Result & "|" & t2.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100|Hello"]);
}
