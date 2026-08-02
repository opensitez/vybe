' vybe-test: vb/vb_version_class_comparison/test_vb_version_to_string_field_count_invalid_throws
' origin: languages/vb/tests/vb/test_vb_version_class_comparison.rs

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
        Dim ver As New Version(1, 2)
        Try
            ver.ToString(3) ' Asking for 3 fields when only 2 exist!
        Catch ex As ArgumentException
            __Check(CStr("ArgumentException Caught on FieldCount Overflow"), "ArgumentException Caught on FieldCount Overflow")
        End Try
    End Sub
End Module
