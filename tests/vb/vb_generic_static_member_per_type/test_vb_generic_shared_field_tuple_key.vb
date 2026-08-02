' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_field_tuple_key
' origin: languages/vb/tests/vb/test_vb_generic_static_member_per_type.rs

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

Class TupleStore(Of T)
    Public Shared StoredTuple As (String, T)
End Class

Module Program
    Sub Main()
        TupleStore(Of Integer).StoredTuple = ("Num", 42)
        TupleStore(Of String).StoredTuple = ("Text", "Val")

        __Check(CStr(TupleStore(Of Integer).StoredTuple.Item2 & "|" & TupleStore(Of String).StoredTuple.Item2), "42|Val")
    End Sub
End Module
