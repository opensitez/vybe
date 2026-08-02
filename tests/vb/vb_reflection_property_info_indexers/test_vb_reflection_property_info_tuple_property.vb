' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_tuple_property
' origin: languages/vb/tests/vb/test_vb_reflection_property_info_indexers.rs

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

Class Entity
    Public Property Pair As (X As Integer, Y As Integer) = (10, 20)
End Class

Module Program
    Sub Main()
        Dim e As New Entity()
        Dim prop = GetType(Entity).GetProperty("Pair")
        Dim tuple As (Integer, Integer) = CType(prop.GetValue(e), (Integer, Integer))
        __Check(CStr(tuple.Item1 & "," & tuple.Item2), "10,20")
    End Sub
End Module
