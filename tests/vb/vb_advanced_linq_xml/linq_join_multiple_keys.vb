' vybe-test: vb/vb_advanced_linq_xml/linq_join_multiple_keys
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

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

Imports System.Linq
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


Class Item
    Public K1 As Integer
    Public K2 As Integer
    Public Val As String
End Class

Module M
    Sub Main()
        Dim arr1 = {New Item With {.K1 = 1, .K2 = 2, .Val = "A"}}
        Dim arr2 = {New Item With {.K1 = 1, .K2 = 2, .Val = "B"}}
        
        Dim query = From a In arr1
                    Join b In arr2 On a.K1 Equals b.K1 And a.K2 Equals b.K2
                    Select a.Val & b.Val
                    
        For Each res In query
            __P(CStr(res))
        Next
        __Check("AB")
    End Sub
End Module
