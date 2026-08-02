' vybe-test: vb/vb_interfaces_basic/interface_property_implementation
' origin: languages/vb/tests/vb/test_vb_interfaces_basic.rs

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

Interface IVehicle
    Property Speed As Integer
End Interface

Class Car
    Implements IVehicle
    
    Private _speed As Integer
    Public Property Speed As Integer Implements IVehicle.Speed
        Get
            Return _speed
        End Get
        Set(value As Integer)
            _speed = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim v As IVehicle = New Car()
        v.Speed = 55
        __Check(CStr(v.Speed), "55")
    End Sub
End Module
