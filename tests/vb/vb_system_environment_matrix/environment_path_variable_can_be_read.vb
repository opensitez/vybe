' vybe-test: vb/vb_system_environment_matrix/environment_path_variable_can_be_read
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
        Dim pathValue As String = Environment.GetEnvironmentVariable("PATH")
        __Check(CStr(pathValue IsNot Nothing), "True")
        __Check(CStr(pathValue.Length > 0), "True")
    End Sub
End Module
