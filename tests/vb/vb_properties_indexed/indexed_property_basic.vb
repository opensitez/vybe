' vybe-test: vb/vb_properties_indexed/indexed_property_basic
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

Class StringGrid
    Private _grid(10) As String
    
    Public Property Cell(index As Integer) As String
        Get
            Return _grid(index)
        End Get
        Set(value As String)
            _grid(index) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim grid As New StringGrid()
        grid.Cell(5) = "Middle"
        grid.Cell(0) = "Start"
        
        __Check(CStr(grid.Cell(0)), "Start")
        __Check(CStr(grid.Cell(5)), "Middle")
    End Sub
End Module
