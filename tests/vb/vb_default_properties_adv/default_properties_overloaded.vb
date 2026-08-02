' vybe-test: vb/vb_default_properties_adv/default_properties_overloaded
' origin: languages/vb/tests/vb/test_vb_default_properties_adv.rs

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

Class Matrix
    Private data(2, 2) As Integer
    
    ' Overloaded Default Properties
    Default Public Property Item(row As Integer, col As Integer) As Integer
        Get
            Return data(row, col)
        End Get
        Set(value As Integer)
            data(row, col) = value
        End Set
    End Property
    
    Default Public Property Item(index As String) As Integer
        Get
            Dim parts = index.Split(","c)
            Return data(Integer.Parse(parts(0)), Integer.Parse(parts(1)))
        End Get
        Set(value As Integer)
            Dim parts = index.Split(","c)
            data(Integer.Parse(parts(0)), Integer.Parse(parts(1))) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim m As New Matrix()
        m(0, 1) = 5
        m("1,2") = 10
        
        __Check(CStr(m(0, 1)), "5")
        __Check(CStr(m("1,2")), "10")
    End Sub
End Module
