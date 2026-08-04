' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_copy_memory_between_pointers
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
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim srcPtr As IntPtr = Marshal.AllocHGlobal(4)
        Dim destPtr As IntPtr = Marshal.AllocHGlobal(4)

        Marshal.WriteInt32(srcPtr, 777)
        Dim tempArr(3) As Byte
        Marshal.Copy(srcPtr, tempArr, 0, 4)
        Marshal.Copy(tempArr, 0, destPtr, 4)

        Dim copiedVal = Marshal.ReadInt32(destPtr)
        Marshal.FreeHGlobal(srcPtr)
        Marshal.FreeHGlobal(destPtr)
        __P(CStr(copiedVal))
        __Check("777")
    End Sub
End Module
