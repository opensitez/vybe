' vybe-test: vb/vb_environment_get_environment_variable/test_vb_environment_get_set_environment_variable_process
' origin: languages/vb/tests/vb/test_vb_environment_get_environment_variable.rs

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
        Environment.SetEnvironmentVariable("VYBE_TEST_VAR", "VybeValue", EnvironmentVariableTarget.Process)
        Dim val = Environment.GetEnvironmentVariable("VYBE_TEST_VAR", EnvironmentVariableTarget.Process)
        __Check(CStr(val), "VybeValue")
    End Sub
End Module
