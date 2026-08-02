' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_struct_method_target
' origin: languages/vb/tests/vb/test_vb_event_delegate_signature_matching.rs

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

Structure Calculator
    Public Factor As Integer
    Public Sub New(f As Integer)
        Factor = f
    End Sub
    Public Function Multiply(val As Integer) As Integer
        Return val * Factor
    End Function
End Structure

Delegate Function MultiplyDel(val As Integer) As Integer

Module Program
    Sub Main()
        Dim c As New Calculator(5)
        Dim d As MultiplyDel = AddressOf c.Multiply
        __Check(CStr(d(10)), "50")
    End Sub
End Module
