' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notoverridable_indexer_property
' origin: languages/vb/tests/vb/test_vb_class_sealed_notinheritable_checks.rs

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

Class BaseContainer
    Default Public Overridable Property Item(index As Integer) As String
        Get
            Return "Base_" & index
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Class FixedContainer
    Inherits BaseContainer
    Default Public NotOverridable Overrides Property Item(index As Integer) As String
        Get
            Return "Fixed_" & index
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim b As BaseContainer = New FixedContainer()
        __Check(CStr(b(5)), "Fixed_5")
    End Sub
End Module
