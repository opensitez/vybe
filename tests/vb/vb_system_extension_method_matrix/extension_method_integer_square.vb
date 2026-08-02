' vybe-test: vb/vb_system_extension_method_matrix/extension_method_integer_square
' origin: languages/vb/tests/vb/test_vb_system_extension_method_matrix.rs

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

Imports System.Runtime.CompilerServices

Module MathExtensions
    <Extension()>
    Public Function Square(value As Integer) As Integer
        Return value * value
    End Function

    <Extension()>
    Public Function DoubleValue(value As Integer) As Integer
        Return value + value
    End Function
End Module

Module M
    Sub Main()
        Dim n As Integer = 7
        __Check(CStr(n.Square()), "49")
        __Check(CStr(n.DoubleValue()), "14")
    End Sub
End Module
