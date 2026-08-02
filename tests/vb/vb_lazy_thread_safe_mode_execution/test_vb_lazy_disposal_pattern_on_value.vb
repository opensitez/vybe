' vybe-test: vb/vb_lazy_thread_safe_mode_execution/test_vb_lazy_disposal_pattern_on_value
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
Imports System.IO

Module Program
    Sub Main()
        Dim lazyStream As New Lazy(Of MemoryStream)(Function() New MemoryStream())
        lazyStream.Value.WriteByte(123)
        __Check(CStr(lazyStream.Value.Length), "1")
        lazyStream.Value.Dispose()
    End Sub
End Module
