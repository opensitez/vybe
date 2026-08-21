' vybe-test: vb/vb_struct_layoutkind/struct_layoutkind
' origin: languages/vb/tests/vb/test_vb_struct_layoutkind.rs

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

Imports System.Runtime.InteropServices
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


<StructLayout(LayoutKind.Explicit)>
Structure UnionType
    <FieldOffset(0)> Public I As Integer
    <FieldOffset(0)> Public B1 As Byte
    <FieldOffset(1)> Public B2 As Byte
    <FieldOffset(2)> Public B3 As Byte
    <FieldOffset(3)> Public B4 As Byte
End Structure

Module M
    Sub Main()
        Dim u As New UnionType()
        u.I = &H12345678
        ' B1 will be the least significant byte on little endian systems (0x78 = 120)
        __P(CStr(u.B1))
        __Check("120")
    End Sub
End Module
