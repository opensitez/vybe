' vybe-test: vb/vb_modules/module_sub_and_function
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
    Sub Greet(name As String)
        __Check(CStr("Hello " & name), "Hello World")
    End Sub
    Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
    Sub Main()
        Greet("World")
        __Check(CStr(Add(3, 4)), "7")
    End Sub
End Module
