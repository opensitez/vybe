' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_constructor_math_constants_precomputation
' origin: languages/vb/tests/vb/test_vb_class_static_shared_constructor.rs

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

Imports System

Class MathLookup
    Public Shared ReadOnly SqrtTwo As Double
    Public Shared ReadOnly SqrtThree As Double
    Shared Sub New()
        SqrtTwo = Math.Sqrt(2.0)
        SqrtThree = Math.Sqrt(3.0)
    End Sub
End Class

Module Program
    Sub Main()
        __Check(CStr(Math.Round(MathLookup.SqrtTwo, 4) & "|" & Math.Round(MathLookup.SqrtThree, 4)), "1.4142|1.7321")
    End Sub
End Module
