' vybe-test: vb/vb_partial_classes/partial_class_fields
' origin: languages/vb/tests/vb/test_vb_partial_classes.rs

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

Partial Class Employee
    Public FirstName As String
End Class

Partial Class Employee
    Public LastName As String
End Class

Module M
    Sub Main()
        Dim e As New Employee()
        e.FirstName = "John"
        e.LastName = "Smith"
        __Check(CStr(e.FirstName & " " & e.LastName), "John Smith")
    End Sub
End Module
