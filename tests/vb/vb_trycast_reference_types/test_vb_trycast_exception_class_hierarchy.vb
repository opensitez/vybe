' vybe-test: vb/vb_trycast_reference_types/test_vb_trycast_exception_class_hierarchy
' origin: languages/vb/tests/vb/test_vb_trycast_reference_types.rs

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
        Dim ex As Exception = New InvalidOperationException("Failed")
        Dim ioe As InvalidOperationException = TryCast(ex, InvalidOperationException)
        Dim ae As ArgumentException = TryCast(ex, ArgumentException)
        __Check(CStr((ioe IsNot Nothing) & "|" & (ae Is Nothing)), "True|True")
    End Sub
End Module
