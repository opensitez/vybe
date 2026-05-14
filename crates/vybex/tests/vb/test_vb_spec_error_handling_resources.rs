use super::helpers::run_vb;

macro_rules! vb_full_spec {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_full_spec!(error_spec_try_catch_handles_simple_exception, r#"Module M
    Sub Main()
        Try
            Throw New Exception("boom")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module"#, ["boom"]);
vb_full_spec!(error_spec_try_catch_handles_specific_exception_type, r#"Module M
    Sub Main()
        Try
            Throw New ArgumentException("bad")
        Catch ex As ArgumentException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module"#, ["bad"]);
vb_full_spec!(error_spec_try_finally_runs_finally_without_exception, r#"Module M
    Sub Main()
        Try
            Console.WriteLine("try")
        Finally
            Console.WriteLine("finally")
        End Try
    End Sub
End Module"#, ["try", "finally"]);
vb_full_spec!(error_spec_try_catch_finally_runs_all_paths, r#"Module M
    Sub Main()
        Try
            Throw New Exception("boom")
        Catch ex As Exception
            Console.WriteLine("catch")
        Finally
            Console.WriteLine("finally")
        End Try
    End Sub
End Module"#, ["catch", "finally"]);
vb_full_spec!(error_spec_throw_new_exception_transfers_control_to_catch, r#"Module M
    Sub Main()
        Try
            Throw New Exception("x")
            Console.WriteLine("after")
        Catch ex As Exception
            Console.WriteLine("caught")
        End Try
    End Sub
End Module"#, ["caught"]);
vb_full_spec!(error_spec_rethrow_preserves_outer_catch_visibility, r#"Module M
    Sub Main()
        Try
            Try
                Throw New Exception("x")
            Catch ex As Exception
                Throw
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module"#, ["x"]);
vb_full_spec!(error_spec_nested_try_inner_catch_handles_exception, r#"Module M
    Sub Main()
        Try
            Try
                Throw New Exception("inner")
            Catch ex As Exception
                Console.WriteLine("caught inner")
            End Try
        Catch ex As Exception
            Console.WriteLine("outer")
        End Try
    End Sub
End Module"#, ["caught inner"]);
vb_full_spec!(error_spec_nested_try_outer_catch_handles_unhandled_inner_exception, r#"Module M
    Sub Main()
        Try
            Try
                Throw New Exception("inner")
            Finally
                Console.WriteLine("inner finally")
            End Try
        Catch ex As Exception
            Console.WriteLine("outer catch")
        End Try
    End Sub
End Module"#, ["inner finally", "outer catch"]);
vb_full_spec!(error_spec_exit_try_skips_remaining_try_body, r#"Module M
    Sub Main()
        Try
            Console.WriteLine("before")
            Exit Try
            Console.WriteLine("after")
        Finally
            Console.WriteLine("finally")
        End Try
    End Sub
End Module"#, ["before", "finally"]);
vb_full_spec!(error_spec_finally_runs_after_return_from_try_block, r#"Module M
    Function Work() As Integer
        Try
            Return 7
        Finally
            Console.WriteLine("finally")
        End Try
    End Function
    Sub Main()
        Console.WriteLine(Work())
    End Sub
End Module"#, ["finally", "7"]);
vb_full_spec!(error_spec_finally_runs_after_return_from_catch_block, r#"Module M
    Function Work() As Integer
        Try
            Throw New Exception("x")
        Catch ex As Exception
            Return 7
        Finally
            Console.WriteLine("finally")
        End Try
    End Function
    Sub Main()
        Console.WriteLine(Work())
    End Sub
End Module"#, ["finally", "7"]);
vb_full_spec!(error_spec_using_disposes_resource_after_scope, r#"Class Probe
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("disposed")
    End Sub
End Class
Module M
    Sub Main()
        Using value As New Probe()
            Console.WriteLine("body")
        End Using
    End Sub
End Module"#, ["body", "disposed"]);
vb_full_spec!(error_spec_using_two_resources_disposes_in_reverse_order, r#"Class Probe
    Implements IDisposable
    Private _name As String
    Public Sub New(name As String)
        _name = name
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine(_name)
    End Sub
End Class
Module M
    Sub Main()
        Using first As New Probe("first"), second As New Probe("second")
            Console.WriteLine("body")
        End Using
    End Sub
End Module"#, ["body", "second", "first"]);
vb_full_spec!(error_spec_using_variable_is_available_inside_scope, r#"Class Probe
    Implements IDisposable
    Public Name As String = "open"
    Public Sub Dispose() Implements IDisposable.Dispose
    End Sub
End Class
Module M
    Sub Main()
        Using value As New Probe()
            Console.WriteLine(value.Name)
        End Using
    End Sub
End Module"#, ["open"]);
vb_full_spec!(error_spec_synclock_allows_mutation_inside_block, r#"Module M
    Sub Main()
        Dim gate As New Object()
        Dim total As Integer = 0
        SyncLock gate
            total += 3
        End SyncLock
        Console.WriteLine(total)
    End Sub
End Module"#, ["3"]);
vb_full_spec!(error_spec_synclock_can_wrap_multiple_statements, r#"Module M
    Sub Main()
        Dim gate As New Object()
        Dim total As Integer = 0
        SyncLock gate
            total += 3
            total += 4
        End SyncLock
        Console.WriteLine(total)
    End Sub
End Module"#, ["7"]);
vb_full_spec!(error_spec_on_error_resume_next_skips_faulting_statement, r#"Module M
    Sub Main()
        On Error Resume Next
        Err.Raise(5)
        Console.WriteLine("after")
    End Sub
End Module"#, ["after"]);
vb_full_spec!(error_spec_on_error_goto_label_jumps_to_handler, r#"Module M
    Sub Main()
        On Error GoTo Handler
        Err.Raise(5)
        Console.WriteLine("after")
        Exit Sub
Handler:
        Console.WriteLine("handled")
    End Sub
End Module"#, ["handled"]);
vb_full_spec!(error_spec_on_error_goto_zero_resets_handler, r#"Module M
    Sub Main()
        On Error Resume Next
        Err.Raise(5)
        On Error GoTo 0
        Console.WriteLine("cleared")
    End Sub
End Module"#, ["cleared"]);
vb_full_spec!(error_spec_err_clear_resets_error_object, r#"Module M
    Sub Main()
        On Error Resume Next
        Err.Raise(5, , "boom")
        Console.WriteLine(Err.Description)
        Err.Clear()
        Console.WriteLine(Err.Number)
    End Sub
End Module"#, ["boom", "0"]);
vb_full_spec!(error_spec_err_description_exposes_current_message, r#"Module M
    Sub Main()
        On Error Resume Next
        Err.Raise(5, , "boom")
        Console.WriteLine(Err.Description)
    End Sub
End Module"#, ["boom"]);
vb_full_spec!(error_spec_catch_when_clause_matches_true_condition, r#"Module M
    Sub Main()
        Try
            Throw New Exception("boom")
        Catch ex As Exception When ex.Message = "boom"
            Console.WriteLine("matched")
        End Try
    End Sub
End Module"#, ["matched"]);
vb_full_spec!(error_spec_catch_when_clause_skips_false_condition, r#"Module M
    Sub Main()
        Try
            Throw New Exception("boom")
        Catch ex As Exception When ex.Message = "other"
            Console.WriteLine("matched")
        Catch ex As Exception
            Console.WriteLine("fallback")
        End Try
    End Sub
End Module"#, ["fallback"]);
vb_full_spec!(error_spec_finally_runs_after_throw_inside_catch, r#"Module M
    Sub Main()
        Try
            Throw New Exception("boom")
        Catch ex As Exception
            Try
                Throw New Exception("other")
            Finally
                Console.WriteLine("finally")
            End Try
        End Try
    End Sub
End Module"#, ["finally"]);
vb_full_spec!(error_spec_using_can_wrap_existing_expression_result, r#"Class Probe
    Implements IDisposable
    Public Shared Function Build() As Probe
        Return New Probe()
    End Function
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("disposed")
    End Sub
End Class
Module M
    Sub Main()
        Using value As Probe = Probe.Build()
            Console.WriteLine("body")
        End Using
    End Sub
End Module"#, ["body", "disposed"]);
vb_full_spec!(error_spec_disposable_class_can_track_dispose_calls, r#"Class Probe
    Implements IDisposable
    Public Shared Count As Integer
    Public Sub Dispose() Implements IDisposable.Dispose
        Count += 1
        Console.WriteLine(Count)
    End Sub
End Class
Module M
    Sub Main()
        Using value As New Probe()
        End Using
    End Sub
End Module"#, ["1"]);
vb_full_spec!(error_spec_nested_using_scopes_dispose_both_resources, r#"Class Probe
    Implements IDisposable
    Private _name As String
    Public Sub New(name As String)
        _name = name
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine(_name)
    End Sub
End Class
Module M
    Sub Main()
        Using outerValue As New Probe("outer")
            Using innerValue As New Probe("inner")
                Console.WriteLine("body")
            End Using
        End Using
    End Sub
End Module"#, ["body", "inner", "outer"]);
vb_full_spec!(error_spec_throw_inside_using_still_disposes_resource, r#"Class Probe
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("disposed")
    End Sub
End Class
Module M
    Sub Main()
        Try
            Using value As New Probe()
                Throw New Exception("boom")
            End Using
        Catch ex As Exception
            Console.WriteLine("caught")
        End Try
    End Sub
End Module"#, ["disposed", "caught"]);
vb_full_spec!(error_spec_synclock_uses_reference_expression, r#"Class Holder
    Public Gate As New Object()
End Class
Module M
    Sub Main()
        Dim holder As New Holder()
        Dim total As Integer = 0
        SyncLock holder.Gate
            total += 8
        End SyncLock
        Console.WriteLine(total)
    End Sub
End Module"#, ["8"]);
vb_full_spec!(error_spec_try_inside_synclock_can_catch_exception, r#"Module M
    Sub Main()
        Dim gate As New Object()
        SyncLock gate
            Try
                Throw New Exception("boom")
            Catch ex As Exception
                Console.WriteLine("caught")
            End Try
        End SyncLock
    End Sub
End Module"#, ["caught"]);
vb_full_spec!(error_spec_using_inside_try_can_dispose_before_catch, r#"Class Probe
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("disposed")
    End Sub
End Class
Module M
    Sub Main()
        Try
            Using value As New Probe()
                Throw New Exception("boom")
            End Using
        Catch ex As Exception
            Console.WriteLine("caught")
        End Try
    End Sub
End Module"#, ["disposed", "caught"]);
vb_full_spec!(error_spec_return_from_using_preserves_return_value_and_disposes, r#"Class Probe
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("disposed")
    End Sub
End Class
Module M
    Function Work() As Integer
        Using value As New Probe()
            Return 9
        End Using
    End Function
    Sub Main()
        Console.WriteLine(Work())
    End Sub
End Module"#, ["disposed", "9"]);
vb_full_spec!(error_spec_return_from_try_finally_preserves_return_value, r#"Module M
    Function Work() As Integer
        Try
            Return 5
        Finally
            Console.WriteLine("finally")
        End Try
    End Function
    Sub Main()
        Console.WriteLine(Work())
    End Sub
End Module"#, ["finally", "5"]);
vb_full_spec!(error_spec_exit_try_enters_finally_block, r#"Module M
    Sub Main()
        Try
            Exit Try
        Finally
            Console.WriteLine("finally")
        End Try
    End Sub
End Module"#, ["finally"]);
vb_full_spec!(error_spec_multiple_catch_blocks_can_select_specific_type, r#"Module M
    Sub Main()
        Try
            Throw New ArgumentException("bad")
        Catch ex As ArgumentException
            Console.WriteLine("arg")
        Catch ex As Exception
            Console.WriteLine("general")
        End Try
    End Sub
End Module"#, ["arg"]);
vb_full_spec!(error_spec_argument_exception_matches_specific_catch_block, r#"Module M
    Sub Main()
        Try
            Throw New ArgumentException("bad")
        Catch ex As ArgumentException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module"#, ["bad"]);
vb_full_spec!(error_spec_application_exception_can_fall_back_to_general_catch, r#"Module M
    Sub Main()
        Try
            Throw New ApplicationException("boom")
        Catch ex As ArgumentException
            Console.WriteLine("arg")
        Catch ex As Exception
            Console.WriteLine("general")
        End Try
    End Sub
End Module"#, ["general"]);
vb_full_spec!(error_spec_finally_can_mutate_local_after_catch, r#"Module M
    Sub Main()
        Dim state As String = "start"
        Try
            Throw New Exception("boom")
        Catch ex As Exception
            state = "catch"
        Finally
            state &= "+finally"
        End Try
        Console.WriteLine(state)
    End Sub
End Module"#, ["catch+finally"]);
vb_full_spec!(error_spec_on_error_resume_next_allows_following_assignment, r#"Module M
    Sub Main()
        Dim value As Integer = 1
        On Error Resume Next
        Err.Raise(5)
        value = 9
        Console.WriteLine(value)
    End Sub
End Module"#, ["9"]);
vb_full_spec!(error_spec_on_error_goto_minus_one_clears_exception_state, r#"Module M
    Sub Main()
        On Error Resume Next
        Err.Raise(5)
        On Error GoTo -1
        Console.WriteLine("cleared")
    End Sub
End Module"#, ["cleared"]);
vb_full_spec!(error_spec_using_supports_interface_typed_resource, r#"Class Probe
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("disposed")
    End Sub
End Class
Module M
    Sub Main()
        Using value As IDisposable = New Probe()
            Console.WriteLine("body")
        End Using
    End Sub
End Module"#, ["body", "disposed"]);
vb_full_spec!(error_spec_try_can_wrap_lambda_invocation, r#"Module M
    Sub Main()
        Dim work As Action = Sub()
            Throw New Exception("boom")
        End Sub
        Try
            work()
        Catch ex As Exception
            Console.WriteLine("caught")
        End Try
    End Sub
End Module"#, ["caught"]);
vb_full_spec!(error_spec_catch_can_inspect_exception_message, r#"Module M
    Sub Main()
        Try
            Throw New Exception("boom")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module"#, ["boom"]);
vb_full_spec!(error_spec_custom_exception_class_can_be_thrown, r#"Class ProblemException
    Inherits Exception
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub
End Class
Module M
    Sub Main()
        Try
            Throw New ProblemException("custom")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module"#, ["custom"]);
vb_full_spec!(error_spec_catch_order_prefers_more_specific_handler, r#"Module M
    Sub Main()
        Try
            Throw New ArgumentException("bad")
        Catch ex As ArgumentException
            Console.WriteLine("specific")
        Catch ex As Exception
            Console.WriteLine("general")
        End Try
    End Sub
End Module"#, ["specific"]);
vb_full_spec!(error_spec_try_without_exception_skips_catch_block, r#"Module M
    Sub Main()
        Try
            Console.WriteLine("ok")
        Catch ex As Exception
            Console.WriteLine("catch")
        End Try
    End Sub
End Module"#, ["ok"]);
vb_full_spec!(error_spec_throw_inside_nested_try_executes_both_finally_blocks, r#"Module M
    Sub Main()
        Try
            Try
                Throw New Exception("boom")
            Finally
                Console.WriteLine("inner")
            End Try
        Catch ex As Exception
            Console.WriteLine("catch")
        Finally
            Console.WriteLine("outer")
        End Try
    End Sub
End Module"#, ["inner", "catch", "outer"]);
vb_full_spec!(error_spec_synclock_inside_using_can_share_same_object, r#"Class Holder
    Implements IDisposable
    Public Gate As New Object()
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("disposed")
    End Sub
End Class
Module M
    Sub Main()
        Using holder As New Holder()
            SyncLock holder.Gate
                Console.WriteLine("locked")
            End SyncLock
        End Using
    End Sub
End Module"#, ["locked", "disposed"]);
