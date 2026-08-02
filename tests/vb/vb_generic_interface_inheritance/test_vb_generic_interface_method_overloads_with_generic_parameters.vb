' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_method_overloads_with_generic_parameters
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Interface ICalculator(Of T)
    Function Add(a As T, b As T) As T
End Interface

Class IntCalculator
    Implements ICalculator(Of Integer)
    Public Function Add(a As Integer, b As Integer) As Integer Implements ICalculator(Of Integer).Add
        Return a + b
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As ICalculator(Of Integer) = New IntCalculator()
        __Check(CStr(calc.Add(15, 25)), "40")
    End Sub
End Module
