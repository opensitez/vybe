' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_override_equals_hashcode
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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

Structure Token(Of T)
    Public Data As T
    Public Sub New(d As T)
        Data = d
    End Sub
    Public Overrides Function Equals(obj As Object) As Boolean
        If Not (TypeOf obj Is Token(Of T)) Then Return False
        Dim other = CType(obj, Token(Of T))
        Return Object.Equals(Data, other.Data)
    End Function
    Public Overrides Function GetHashCode() As Integer
        If Data Is Nothing Then Return 0
        Return Data.GetHashCode()
    End Function
End Structure

Module Program
    Sub Main()
        Dim t1 As New Token(Of String)("ABC")
        Dim t2 As New Token(Of String)("ABC")
        Dim t3 As New Token(Of String)("XYZ")
        __P(CStr(t1.Equals(t2) & "|" & t1.Equals(t3)))
        __Check("True|False")
    End Sub
End Module
