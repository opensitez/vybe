' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_date_array
' origin: languages/vb/tests/vb/test_vb_array_resize_preserve_semantics.rs

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
        Dim dates(0) As DateTime
        dates(0) = New DateTime(2025, 1, 1)
        ReDim Preserve dates(1)
        __Check(CStr(dates(0).ToString("yyyy-MM-dd")), "2025-01-01")
        __Check(CStr(dates(1) = DateTime.MinValue), "True")
    End Sub
End Module
