' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_property_conflict_disambiguation
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface IAlpha
    ReadOnly Property Value As Integer
End Interface

Interface IBeta
    ReadOnly Property Value As String
End Interface

Class Component
    Implements IAlpha, IBeta
    Public ReadOnly Property AlphaValue As Integer Implements IAlpha.Value
        Get
            Return 42
        End Get
    End Property
    Public ReadOnly Property BetaValue As String Implements IBeta.Value
        Get
            Return "FortyTwo"
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim c As New Component()
        Dim a As IAlpha = c
        Dim b As IBeta = c
        __Check(CStr(a.Value & "|" & b.Value), "42|FortyTwo")
    End Sub
End Module
