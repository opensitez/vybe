' vybe-test: vb/vb_comprehensive/class_bare_method_call_resolves_to_me
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
    Class Calc
        Public Value As Integer = 0

        Sub Add(n As Integer)
            Me.Value = Me.Value + n
        End Sub

        Sub AddTwice(n As Integer)
            Add(n)
            Add(n)
        End Sub
    End Class

    Sub Main()
        Dim c As New Calc()
        c.AddTwice(5)
        __Check(CStr(c.Value), "10")
    End Sub
End Module
