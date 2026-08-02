' vybe-test: vb/vb_comprehensive/with_statement
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

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
    Class Person
        Public Name As String
        Public Age As Integer
    End Class

    Sub Main()
        Dim p As New Person()
        p.Name = "Alice"
        p.Age = 25
        __Check(CStr(p.Name & " is " & CStr(p.Age)), "Alice is 25")
    End Sub
End Module
