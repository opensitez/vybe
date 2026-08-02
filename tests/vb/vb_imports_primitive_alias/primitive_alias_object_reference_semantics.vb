' vybe-test: vb/vb_imports_primitive_alias/primitive_alias_object_reference_semantics
' origin: languages/vb/tests/vb/test_vb_imports_primitive_alias.rs

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

Imports ObjAlias = System.Object

Module M
    Class Holder
        Public Value As String = "A"
    End Class

    Sub Main()
        Dim left As ObjAlias = New Holder()
        Dim right As ObjAlias = left
        Dim holder As Holder = CType(right, Holder)
        holder.Value = "B"
        __Check(CStr(CType(left, Holder).Value), "B")
    End Sub
End Module
