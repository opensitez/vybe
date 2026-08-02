' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_property_per_type_isolation
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

Class PropertyHolder(Of T)
    Private Shared _data As String
    Public Shared Property Data As String
        Get
            Return _data
        End Get
        Set(value As String)
            _data = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        PropertyHolder(Of Integer).Data = "IntData"
        PropertyHolder(Of Double).Data = "DoubleData"
        __Check(CStr(PropertyHolder(Of Integer).Data & "|" & PropertyHolder(Of Double).Data), "IntData|DoubleData")
    End Sub
End Module
