' vybe-test: vb/vb_oop_edges/constructor_chaining_mybase_new
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

Class Base
    Public Val As Integer
    Public Sub New(v As Integer)
        Val = v
    End Sub
End Class

Class Derived
    Inherits Base
    Public Sub New()
        MyBase.New(20)
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        __Check(CStr(d.Val), "20")
    End Sub
End Module
