' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_constructor_exception_propagation
' origin: languages/vb/tests/vb/test_vb_class_static_shared_constructor.rs

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

Class FailingStatic
    Shared Sub New()
        Throw New InvalidOperationException("SharedInitFailed")
    End Sub
    Public Shared Sub Touch()
    End Sub
End Class

Module Program
    Sub Main()
        Try
            FailingStatic.Touch()
        Catch ex As Exception
            __Check(CStr(ex.GetType().Name), "TypeInitializationException")
        End Try
    End Sub
End Module
