' vybe-test: vb/vb_inheritance/t09_constructor_chain_three_levels
' origin: languages/vb/tests/vb/vb_inheritance_test.rs

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

Class Level1
    Public Tag As String

    Sub New()
        Tag = "L1"
        __Check(CStr("Level1.New"), "Level1.New")
    End Sub
End Class

Class Level2
    Inherits Level1

    Sub New()
        MyBase.New()
        Tag = Tag & "-L2"
        __Check(CStr("Level2.New"), "Level2.New")
    End Sub
End Class

Class Level3
    Inherits Level2

    Sub New()
        MyBase.New()
        Tag = Tag & "-L3"
        __Check(CStr("Level3.New"), "Level3.New")
    End Sub
End Class

Dim x As New Level3()
__Check(CStr(x.Tag), "L1-L2-L3")
