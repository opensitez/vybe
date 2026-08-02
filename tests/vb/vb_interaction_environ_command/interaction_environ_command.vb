' vybe-test: vb/vb_interaction_environ_command/interaction_environ_command
' origin: languages/vb/tests/vb/test_vb_interaction_environ_command.rs

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

Module M
    Sub Main()
        ' Environ gets an environment variable
        Dim pathVar = Environ("PATH")
        __Check(CStr(pathVar IsNot Nothing), "True")
        
        ' Command gets the command line arguments as a string
        Dim cmd = Command()
        __Check(CStr(cmd IsNot Nothing), "True")
    End Sub
End Module
