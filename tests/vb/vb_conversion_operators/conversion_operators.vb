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
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
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
        __P(CStr(num))
        
        ' Explicit conversion from Integer to Digit
        Dim d2 As Digit = CType(23, Digit)
        __P(CStr(d2.Value))
        __Check("5
3")
    End Sub
End Module
