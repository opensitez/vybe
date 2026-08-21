' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_custom_attribute
' origin: languages/vb/tests/vb/test_vb_reflection_field_info_get_set.rs

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

Imports System
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


<AttributeUsage(AttributeTargets.Field)>
Class RangeAttribute
    Inherits Attribute
    Public Min As Integer
    Public Max As Integer
    Public Sub New(min As Integer, max As Integer)
        Me.Min = min
        Me.Max = max
    End Sub
End Class

Class Form
    <Range(1, 100)>
    Public Percentage As Integer
End Class

Module Program
    Sub Main()
        Dim field = GetType(Form).GetField("Percentage")
        Dim attr = CType(field.GetCustomAttributes(GetType(RangeAttribute), False)(0), RangeAttribute)
        __P(CStr(attr.Min & " To " & attr.Max))
        __Check("1 To 100")
    End Sub
End Module
