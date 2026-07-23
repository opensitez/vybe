use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Destructor (Finalize) & IDisposable Pattern
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_idisposable_pattern_dispose_call() {
    let src = r#"
Imports System

Class ResourceHandler
    Implements IDisposable

    Public IsDisposed As Boolean = False

    Public Sub Dispose() Implements IDisposable.Dispose
        IsDisposed = True
        Console.WriteLine("Disposed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim rh As New ResourceHandler()
        rh.Dispose()
        Console.WriteLine(rh.IsDisposed)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Disposed", "True"]);
}

#[test]
fn test_vb_idisposable_using_statement_scope() {
    let src = r#"
Imports System

Class ManagedBuffer
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Buffer Disposed")
    End Sub
End Class

Module Program
    Sub Main()
        Using buf As New ManagedBuffer()
            Console.WriteLine("Inside Using")
        End Using
        Console.WriteLine("Outside Using")
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Inside Using", "Buffer Disposed", "Outside Using"]
    );
}

#[test]
fn test_vb_standard_disposable_pattern_suppress_finalize() {
    let src = r#"
Imports System

Class DisposableResource
    Implements IDisposable

    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                Console.WriteLine("Disposing managed")
            End If
            disposedValue = True
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Module Program
    Sub Main()
        Using res As New DisposableResource()
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Disposing managed"]);
}

#[test]
fn test_vb_finalize_method_override() {
    let src = r#"
Imports System

Class FinalizableClass
    Protected Overrides Sub Finalize()
        Try
            Console.WriteLine("Finalized")
        Finally
            MyBase.Finalize()
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim fc As New FinalizableClass()
        fc = Nothing
        Console.WriteLine("Done")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Done"]);
}
