' vybe-test: vb/vb_system_environment_variables_matrix/environment_variables_expand_tokens_and_command_line
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
        Dim expanded As String = Environment.ExpandEnvironmentVariables("prefix_%PATH%")
        __Check(CStr(expanded.Contains("prefix_")), "True")
        Dim args() As String = Environment.GetCommandLineArgs()
        __Check(CStr(args.Length >= 1), "True")
    End Sub
End Module
