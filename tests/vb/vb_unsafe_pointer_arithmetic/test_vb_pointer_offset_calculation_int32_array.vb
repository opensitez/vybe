' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_offset_calculation_int32_array
' origin: languages/vb/tests/vb/test_vb_unsafe_pointer_arithmetic.rs

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


Module Program
    Sub Main()
        Dim numbers As Integer() = {100, 200, 300, 400}
        Dim handle = GCHandle.Alloc(numbers, GCHandleType.Pinned)
        Dim baseAddr = handle.AddrOfPinnedObject()

        ' Offset for element index 2: 2 * Marshal.SizeOf(GetType(Integer)) = 8 bytes
        Dim elem2Addr = IntPtr.Add(baseAddr, 2 * 4)
        Dim val = Marshal.ReadInt32(elem2Addr)
        handle.Free()
        __P(CStr(val))
        __Check("300")
    End Sub
End Module
