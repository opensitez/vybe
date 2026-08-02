' vybe-test: vb/vb_local_constants_adv/local_constants_adv
' origin: languages/vb/tests/vb/test_vb_local_constants_adv.rs

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
        ' Const inside a method
        Const MaxLimit As Integer = 100
        Const Greeting As String = "Hello"
        
        __Check(CStr(MaxLimit), "100")
        __Check(CStr(Greeting), "Hello")
    End Sub
End Module
