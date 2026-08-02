' vybe-test: vb/vb_environment_command_line_args/test_vb_environment_os_version
' origin: languages/vb/tests/vb/test_vb_environment_command_line_args.rs

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
        Dim os = Environment.OSVersion
        __Check(CStr(os.Platform.ToString() & "|" & os.Version.Major > 0), "Unix|True")
    End Sub
End Module
