' vybe-test: vb/vb_system_environment_matrix/environment_user_and_machine_names_are_populated
' origin: languages/vb/tests/vb/test_vb_system_environment_matrix.rs

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

Module M
    Sub Main()
        __Check(CStr(Not String.IsNullOrWhiteSpace(Environment.UserName)), "True")
        __Check(CStr(Not String.IsNullOrWhiteSpace(Environment.MachineName)), "True")
        __Check(CStr(Not String.IsNullOrWhiteSpace(Environment.OSVersion.ToString())), "True")
    End Sub
End Module
