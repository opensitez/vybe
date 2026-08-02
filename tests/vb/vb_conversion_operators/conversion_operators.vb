' vybe-test: vb/vb_conversion_operators/conversion_operators
' origin: languages/vb/tests/vb/test_vb_conversion_operators.rs

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

Structure Digit
    Public Value As Byte
    
    Public Sub New(val As Byte)
        Value = val
    End Sub
    
    ' Widening (Implicit) conversion
    Public Shared Widening Operator CType(d As Digit) As Integer
        Return CInt(d.Value)
    End Operator
    
    ' Narrowing (Explicit) conversion
    Public Shared Narrowing Operator CType(i As Integer) As Digit
        Return New Digit(CByte(i Mod 10))
    End Operator
End Structure

Module M
    Sub Main()
        Dim d As New Digit(5)
        
        ' Implicit conversion to Integer
        Dim num As Integer = d
        __Check(CStr(num), "5")
        
        ' Explicit conversion from Integer to Digit
        Dim d2 As Digit = CType(23, Digit)
        __Check(CStr(d2.Value), "3")
    End Sub
End Module
