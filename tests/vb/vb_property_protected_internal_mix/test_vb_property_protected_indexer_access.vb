' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_protected_indexer_access
' origin: languages/vb/tests/vb/test_vb_property_protected_internal_mix.rs

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

Class BaseArray
    Private arr(2) As Integer
    Protected Default Property Item(idx As Integer) As Integer
        Get
            Return arr(idx)
        End Get
        Set(value As Integer)
            arr(idx) = value
        End Set
    End Property
End Class

Class CustomArray
    Inherits BaseArray
    Public Sub SetValue(idx As Integer, val As Integer)
        MyBase.Item(idx) = val
    End Sub
    Public Function GetValue(idx As Integer) As Integer
        Return MyBase.Item(idx)
    End Function
End Class

Module Program
    Sub Main()
        Dim ca As New CustomArray()
        ca.SetValue(0, 99)
        __Check(CStr(ca.GetValue(0)), "99")
    End Sub
End Module
