' vybe-test: vb/vb_custom_attributes_advanced/custom_attributes_named_arguments
' origin: languages/vb/tests/vb/test_vb_custom_attributes_advanced.rs

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


<AttributeUsage(AttributeTargets.Class Or AttributeTargets.Method)>
Public Class InfoAttribute
    Inherits Attribute

    Public Property Author As String
        Public Property Version As String

            Public Sub New()
            End Sub
        End Class

        <Info(Author:="Alice", Version:="1.0")>
        Class MyComponent
        End Class

        Module M
            Sub Main()
                Dim t As Type = GetType(MyComponent)
                Dim attrs() As Object = t.GetCustomAttributes(GetType(InfoAttribute), False)
                Dim info As InfoAttribute = DirectCast(attrs(0), InfoAttribute)

                __P(CStr(info.Author))
                __P(CStr(info.Version))
                __Check("Alice
1.0")
            End Sub
        End Module
