' vybe-test: vb/vb_string_builder_replace_insert_remove/test_vb_string_builder_exceed_max_capacity_throws
' origin: languages/vb/tests/vb/test_vb_string_builder_replace_insert_remove.rs

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
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder(5, 5)
        Try
            sb.Append("ExceedMaxCapacity")
        Catch ex As ArgumentOutOfRangeException
            __Check(CStr("ArgumentOutOfRangeException Caught"), "ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
