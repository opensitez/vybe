' vybe-test: vb/vb_generic_constraint_class/generic_constraint_class
' origin: languages/vb/tests/vb/test_vb_generic_constraint_class.rs

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

' Generic constraint As Class requires T to be a reference type
Class ReferenceCache(Of T As Class)
    Public Property Item As T
End Class

Module M
    Sub Main()
        Dim c As New ReferenceCache(Of String)()
        c.Item = "Hello"
        __Check(CStr(c.Item), "Hello")
    End Sub
End Module
