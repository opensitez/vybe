' vybe-test: vb/vb_system_process_matrix/process_lookup_missing_id_throws
' origin: languages/vb/tests/vb/test_vb_system_process_matrix.rs

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

Imports System.Diagnostics

Module M
    Sub Main()
        Dim threw As Boolean = False

        Try
            Process.GetProcessById(-1)
        Catch ex As Exception
            threw = True
        End Try

        __Check(CStr(threw), "True")
    End Sub
End Module
