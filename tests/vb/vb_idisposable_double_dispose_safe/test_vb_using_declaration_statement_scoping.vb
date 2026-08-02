' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_using_declaration_statement_scoping
' origin: languages/vb/tests/vb/test_vb_idisposable_double_dispose_safe.rs

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

Class ScopeTracker
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        __Check(CStr("ScopeTracker Disposed"), "Doing Work in Inner Scope")
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Using res As New ScopeTracker()
                __Check(CStr("Doing Work in Inner Scope"), "ScopeTracker Disposed")
            End Using
        End Sub()
        __Check(CStr("Outer Scope"), "Outer Scope")
    End Sub
End Module
