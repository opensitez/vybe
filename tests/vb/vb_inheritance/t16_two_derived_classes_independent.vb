' vybe-test: vb/vb_inheritance/t16_two_derived_classes_independent
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

Class Base
    Public Val As Integer = 0
End Class

Class ChildA
    Inherits Base

    Sub New()
        MyBase.New()
        Val = 1
    End Sub
End Class

Class ChildB
    Inherits Base

    Sub New()
        MyBase.New()
        Val = 2
    End Sub
End Class

Dim a As New ChildA()
Dim b As New ChildB()
__Check(CStr(a.Val), "1")
__Check(CStr(b.Val), "2")
