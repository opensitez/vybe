' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_from_base64_span_buffer
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
        ' ⛔ VB cannot DECLARE a `Span` local — it is a ref struct (BC30668)
        ' — but an array passes straight to a `Span` PARAMETER, so the span
        ' overload is still the one called.
        Dim span = "AQID".ToCharArray()
        Dim dest(3) As Byte
        Dim bytesWritten As Integer
        Dim ok = Convert.TryFromBase64Chars(span, dest, bytesWritten)
        ' ⛔ `String.Join(sep, arr, index, count)` takes a STRING array — with a
        ' Byte() it binds to the `ParamArray Object()` form and renders
        ' "System.Byte[],0,3". Copy the written prefix out first.
        Dim written(bytesWritten - 1) As Byte
        Array.Copy(dest, written, bytesWritten)
        __P(CStr(ok) & ":" & String.Join(",", written))
        __Check("True:1,2,3")
    End Sub
End Module
