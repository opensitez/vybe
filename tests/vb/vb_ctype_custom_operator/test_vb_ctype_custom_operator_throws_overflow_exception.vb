' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_custom_operator_throws_overflow_exception
' origin: languages/vb/tests/vb/test_vb_ctype_custom_operator.rs

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

Class BoundedVal
    Public Value As Long
    Public Sub New(v As Long)
        Value = v
    End Sub

    Public Shared Narrowing Operator CType(b As BoundedVal) As Byte
        If b.Value < 0 OrElse b.Value > 255 Then Throw New OverflowException("BoundedVal Byte Overflow")
        Return CByte(b.Value)
    End Shared Narrowing Operator
End Class

Module Program
    Sub Main()
        Dim bv As New BoundedVal(1000)
        Try
            Dim b As Byte = CType(bv, Byte)
        Catch ex As OverflowException
            __Check(CStr(ex.Message), "BoundedVal Byte Overflow")
        End Try
    End Sub
End Module
