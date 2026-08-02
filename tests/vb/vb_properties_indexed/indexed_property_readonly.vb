' vybe-test: vb/vb_properties_indexed/indexed_property_readonly
' origin: languages/vb/tests/vb/test_vb_properties_indexed.rs

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

Class MathTable
    Public ReadOnly Property Multiplier(factor As Integer) As Integer
        Get
            Return factor * 10
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim table As New MathTable()
        __Check(CStr(table.Multiplier(5)), "50")
        __Check(CStr(table.Multiplier(9)), "90")
    End Sub
End Module
