' vybe-test: vb/vb_class/class_shared_method
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
    Class MathHelper
        Shared Function Add(a As Double, b As Double) As Double
            Add = a + b
        End Function

        Shared Function Multiply(a As Double, b As Double) As Double
            Multiply = a * b
        End Function
    End Class

    Sub Main()
        __Check(CStr(MathHelper.Add(3, 4)), "7")
        __Check(CStr(MathHelper.Multiply(5, 6)), "30")
    End Sub
End Module
