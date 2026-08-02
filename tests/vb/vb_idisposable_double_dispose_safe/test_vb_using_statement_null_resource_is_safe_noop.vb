' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_using_statement_null_resource_is_safe_noop
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

Module Program
    Sub Main()
        Dim res As IDisposable = Nothing
        Using res
            __Check(CStr("Inside Null Using"), "Inside Null Using")
        End Using
        __Check(CStr("After Null Using"), "After Null Using")
    End Sub
End Module
