' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_overloaded_indexed_properties
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

Class MultiIndexer
    Default Public Property Item(i As Integer) As String
        Get : Return "Int_" & i : End Get
        Set(value As String) : End Set
    End Property
    Default Public Property Item(s As String) As String
        Get : Return "Str_" & s : End Get
        Set(value As String) : End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim mi As New MultiIndexer()
        Dim pInt = GetType(MultiIndexer).GetProperty("Item", {GetType(Integer)})
        Dim pStr = GetType(MultiIndexer).GetProperty("Item", {GetType(String)})

        __Check(CStr(pInt.GetValue(mi, {5}) & "|" & pStr.GetValue(mi, {"abc"})), "Int_5|Str_abc")
    End Sub
End Module
