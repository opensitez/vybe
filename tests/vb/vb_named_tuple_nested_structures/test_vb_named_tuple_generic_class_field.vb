' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_generic_class_field
' origin: languages/vb/tests/vb/test_vb_named_tuple_nested_structures.rs

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

Class StateHolder(Of T)
    Public CurrentState As (Status As String, Data As T)
    Public Sub New(s As String, d As T)
        CurrentState = (s, d)
    End Sub
End Class

Module Program
    Sub Main()
        Dim sh As New StateHolder(Of Integer)("OK", 200)
        __Check(CStr(sh.CurrentState.Status & "=" & sh.CurrentState.Data), "OK=200")
    End Sub
End Module
