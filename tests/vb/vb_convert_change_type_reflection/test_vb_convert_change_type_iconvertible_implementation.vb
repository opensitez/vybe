' vybe-test: vb/vb_convert_change_type_reflection/test_vb_convert_change_type_iconvertible_implementation
' origin: languages/vb/tests/vb/test_vb_convert_change_type_reflection.rs

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

Imports System
Imports System.Globalization

Class CustomConvertible
    Implements IConvertible

    Public Value As Integer = 777

    Public Function ToInt32(provider As IFormatProvider) As Integer Implements IConvertible.ToInt32
        Return Value
    End Function

    Public Function GetTypeCode() As TypeCode Implements IConvertible.GetTypeCode
        Return TypeCode.Object
    End Function
    Public Function ToBoolean(provider As IFormatProvider) As Boolean Implements IConvertible.ToBoolean
        Return Value <> 0
    End Function
    Public Function ToByte(provider As IFormatProvider) As Byte Implements IConvertible.ToByte
        Return CByte(Value)
    End Function
    Public Function ToChar(provider As IFormatProvider) As Char Implements IConvertible.ToChar
        Return "X"c
    End Function
    Public Function ToDateTime(provider As IFormatProvider) As DateTime Implements IConvertible.ToDateTime
        Return DateTime.MinValue
    End Function
    Public Function ToDecimal(provider As IFormatProvider) As Decimal Implements IConvertible.ToDecimal
        Return Value
    End Function
    Public Function ToDouble(provider As IFormatProvider) As Double Implements IConvertible.ToDouble
        Return Value
    End Function
    Public Function ToInt16(provider As IFormatProvider) As Short Implements IConvertible.ToInt16
        Return CShort(Value)
    End Function
    Public Function ToInt64(provider As IFormatProvider) As Long Implements IConvertible.ToInt64
        Return Value
    End Function
    Public Function ToSByte(provider As IFormatProvider) As SByte Implements IConvertible.ToSByte
        Return CSByte(Value)
    End Function
    Public Function ToSingle(provider As IFormatProvider) As Single Implements IConvertible.ToSingle
        Return Value
    End Function
    Public Function ToString(provider As IFormatProvider) As String Implements IConvertible.ToString
        Return Value.ToString()
    End Function
    Public Function ToType(conversionType As Type, provider As IFormatProvider) As Object Implements IConvertible.ToType
        If conversionType Is GetType(Integer) Then Return Value
        Throw New InvalidCastException()
    End Function
    Public Function ToUInt16(provider As IFormatProvider) As UShort Implements IConvertible.ToUInt16
        Return CUShort(Value)
    End Function
    Public Function ToUInt32(provider As IFormatProvider) As UInteger Implements IConvertible.ToUInt32
        Return CUInt(Value)
    End Function
    Public Function ToUInt64(provider As IFormatProvider) As ULong Implements IConvertible.ToUInt64
        Return CULng(Value)
    End Function
End Class

Module Program
    Sub Main()
        Dim cc As New CustomConvertible()
        Dim num As Object = Convert.ChangeType(cc, GetType(Integer))
        __P(CStr(num))
        __Check("777")
    End Sub
End Module
