' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_single_class_multiple_methods_same_interface
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

Interface IMathOps
    Function Add(a As Integer, b As Integer) As Integer
    Function Add(a As Double, b As Double) As Double
End Interface

Class Calculator
    Implements IMathOps
    Public Function Add(a As Integer, b As Integer) As Integer Implements IMathOps.Add
        Return a + b
    End Function
    Public Function Add(a As Double, b As Double) As Double Implements IMathOps.Add
        Return a + b
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As IMathOps = New Calculator()
        __Check(CStr(calc.Add(5, 10) & "|" & calc.Add(2.5, 3.5)), "15|6")
    End Sub
End Module
