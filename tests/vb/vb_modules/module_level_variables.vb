' vybe-test: vb/vb_modules/module_level_variables
' origin: languages/vb/tests/vb/test_vb_modules.rs

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
    Dim counter As Integer = 0
    Sub Increment()
        counter = counter + 1
    End Sub
    Sub Main()
        Increment()
        Increment()
        Increment()
        __Check(CStr(counter), "3")
    End Sub
End Module
