' vybe-test: vb/vb_oop_edges/partial_class_with_implements
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

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

Interface I1
    Sub M1()
End Interface

Interface I2
    Sub M2()
End Interface

Partial Class C
    Implements I1
    Public Sub M1() Implements I1.M1
        __Check(CStr("M1"), "M1")
    End Sub
End Class

Partial Class C
    Implements I2
    Public Sub M2() Implements I2.M2
        __Check(CStr("M2"), "M2")
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New C()
        c.M1()
        c.M2()
    End Sub
End Module
