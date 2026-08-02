' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_two_type_parameters
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Interface IMapping(Of TKey, TValue)
    Function Map(key As TKey) As TValue
End Interface

Class IntToStringMapper
    Implements IMapping(Of Integer, String)
    Public Function Map(key As Integer) As String Implements IMapping(Of Integer, String).Map
        Return "Value_" & key
    End Function
End Class

Module Program
    Sub Main()
        Dim m As IMapping(Of Integer, String) = New IntToStringMapper()
        __Check(CStr(m.Map(42)), "Value_42")
    End Sub
End Module
