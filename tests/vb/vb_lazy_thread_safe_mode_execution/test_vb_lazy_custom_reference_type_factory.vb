' vybe-test: vb/vb_lazy_thread_safe_mode_execution/test_vb_lazy_custom_reference_type_factory
' origin: languages/vb/tests/vb/test_vb_lazy_thread_safe_mode_execution.rs

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

Class Config
    Public Property Port As Integer
End Class

Module Program
    Sub Main()
        Dim lazyConfig As New Lazy(Of Config)(Function() New Config With {.Port = 8080})
        __Check(CStr(lazyConfig.Value.Port), "8080")
    End Sub
End Module
