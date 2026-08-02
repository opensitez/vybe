' vybe-test: vb/vb_oop_edges/interface_property_implementation
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

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

Interface IData
    Property Value As Integer
End Interface

Class Data
    Implements IData
    Private _val As Integer
    
    Public Property Value As Integer Implements IData.Value
        Get
            Return _val
        End Get
        Set(v As Integer)
            _val = v
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim d As IData = New Data()
        d.Value = 42
        __Check(CStr(d.Value), "42")
    End Sub
End Module
