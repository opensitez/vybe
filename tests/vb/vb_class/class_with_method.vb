' vybe-test: vb/vb_class/class_with_method
' origin: languages/vb/tests/vb/vb_class_test.rs

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

Module Program
    Class Calculator
        Public Result As Double

        Sub New()
            Me.Result = 0
        End Sub

        Function Add(a As Double, b As Double) As Double
            Add = a + b
        End Function

        Sub AddToResult(value As Double)
            Me.Result = Me.Result + value
        End Sub
    End Class

    Sub Main()
        Dim calc As New Calculator()
        __Check(CStr(calc.Add(3, 4)), "7")
        calc.AddToResult(10)
        calc.AddToResult(20)
        __Check(CStr(calc.Result), "30")
    End Sub
End Module
