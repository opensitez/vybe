' vybe-test: vb/vb_properties_default_advanced/default_properties_multiple_parameters
' origin: languages/vb/tests/vb/test_vb_properties_default_advanced.rs

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
    
    ' Default property can have multiple parameters
    Default Public Property Item(row As Integer, col As Integer) As Integer
        Get
            Return data(row, col)
        End Get
        Set(value As Integer)
            data(row, col) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim m As New Matrix()
        
        ' Calling Default Property with multiple arguments
        m(1, 2) = 42
        __Check(CStr(m(1, 2)), "42")
    End Sub
End Module
