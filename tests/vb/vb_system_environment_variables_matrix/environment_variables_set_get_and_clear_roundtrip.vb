' vybe-test: vb/vb_system_environment_variables_matrix/environment_variables_set_get_and_clear_roundtrip
' origin: languages/vb/tests/vb/test_vb_system_environment_variables_matrix.rs

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
        Dim key As String = "VYBE_VB_ENV_MATRIX"
        Environment.SetEnvironmentVariable(key, "present")
        __Check(CStr(Environment.GetEnvironmentVariable(key) = "present"), "True")

        Environment.SetEnvironmentVariable(key, Nothing)
        __Check(CStr(Environment.GetEnvironmentVariable(key) Is Nothing), "True")
    End Sub
End Module
