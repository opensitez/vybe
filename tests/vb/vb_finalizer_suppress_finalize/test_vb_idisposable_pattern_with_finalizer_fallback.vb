' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_idisposable_pattern_with_finalizer_fallback
' origin: languages/vb/tests/vb/test_vb_finalizer_suppress_finalize.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Class FullDisposablePattern
    Implements IDisposable

    Public Property CleanedFromDispose As Boolean = False
    Public Property CleanedFromFinalizer As Boolean = False
    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                CleanedFromDispose = True
            Else
                CleanedFromFinalizer = True
            End If
            disposedValue = True
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(disposing:=False)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New FullDisposablePattern()
        obj.Dispose()
        __Check(CStr(obj.CleanedFromDispose & "|" & obj.CleanedFromFinalizer), "True|False")
    End Sub
End Module
