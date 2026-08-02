' vybe-test: vb/vb_system_environment_variables_matrix/environment_variables_map_contains_common_keys
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
        Dim values As Collections.IDictionary = Environment.GetEnvironmentVariables()
        __Check(CStr(values IsNot Nothing), "True")
        __Check(CStr(values.Count > 0), "True")
    End Sub
End Module
