' vybe-test: vb/vb_oop_edges/constructor_chaining_myclass_new
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

Class C
    Public Val As Integer
    
    Public Sub New()
        MyClass.New(10)
    End Sub
    
    Public Sub New(v As Integer)
        Val = v
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New C()
        __Check(CStr(c.Val), "10")
    End Sub
End Module
