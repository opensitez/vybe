' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_read_only_property
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IIdentifiable
    ReadOnly Property Id As Integer
End Interface

Class Item
    Implements IIdentifiable
    Private _id As Integer
    Public Sub New(id As Integer)
        _id = id
    End Sub
    Public ReadOnly Property Id As Integer Implements IIdentifiable.Id
        Get
            Return _id
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim item As IIdentifiable = New Item(42)
        __Check(CStr(item.Id), "42")
    End Sub
End Module
