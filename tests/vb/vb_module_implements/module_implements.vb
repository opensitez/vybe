' vybe-test: vb/vb_module_implements/module_implements
' origin: languages/vb/tests/vb/test_vb_module_implements.rs

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

Interface IRunnable
    Sub Run()
End Interface

' Modules cannot implement interfaces in standard VB.NET.
' We wrap this in a class to test the syntax for Implements instead.
Class Runner
    Implements IRunnable
    
    Public Sub Run() Implements IRunnable.Run
        __Check(CStr("Running"), "Running")
    End Sub
End Class

Module M
    Sub Main()
        Dim r As IRunnable = New Runner()
        r.Run()
    End Sub
End Module
