' vybe-test: vb/vb_system_environment_variables_matrix/environment_variable_target_machine_and_process_set_supported
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
        Dim key As String = "VYBE_VB_ENV_MACHINE_TEST"
        Environment.SetEnvironmentVariable(key, "1", EnvironmentVariableTarget.Process)

        Dim processValue As String = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Process)
        __Check(CStr(processValue = "1"), "True")

        Dim machineValue As String = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine)
        __Check(CStr(Not String.IsNullOrWhiteSpace(machineValue)), "True")
    End Sub
End Module
