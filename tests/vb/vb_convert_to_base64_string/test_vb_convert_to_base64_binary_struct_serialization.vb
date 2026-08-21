' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_to_base64_binary_struct_serialization
' origin: languages/vb/tests/vb/test_vb_convert_to_base64_string.rs

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


<StructLayout(LayoutKind.Sequential)>
Structure RecordHeader
    Public Magic As Integer
    Public Length As Short
End Structure

Module Program
    Sub Main()
        Dim h As New RecordHeader With {.Magic = &H41424344, .Length = 100}
        Dim size = Marshal.SizeOf(GetType(RecordHeader))
        Dim ptr = Marshal.AllocHGlobal(size)
        Marshal.StructureToPtr(h, ptr, False)

        Dim bytes(size - 1) As Byte
        Marshal.Copy(ptr, bytes, 0, size)
        Marshal.FreeHGlobal(ptr)

        Dim b64 = Convert.ToBase64String(bytes)
        __P(CStr(b64.Length > 0))
        __Check("True")
    End Sub
End Module
