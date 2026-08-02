' vybe-test: vb/vb_destructor_finalizer_pattern/test_vb_idisposable_using_statement_scope
' origin: languages/vb/tests/vb/test_vb_destructor_finalizer_pattern.rs

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

Class ManagedBuffer
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        __Check(CStr("Buffer Disposed"), "Inside Using")
    End Sub
End Class

Module Program
    Sub Main()
        Using buf As New ManagedBuffer()
            __Check(CStr("Inside Using"), "Buffer Disposed")
        End Using
        __Check(CStr("Outside Using"), "Outside Using")
    End Sub
End Module
