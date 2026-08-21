' vybe-test: vb/vb_linq_distinct_custom_equality_comparer/test_vb_linq_distinct_custom_iequalitycomparer
' origin: languages/vb/tests/vb/test_vb_linq_distinct_custom_equality_comparer.rs

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

Imports System.Collections.Generic
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


Class Product
    Public Property ID As Integer
        Public Property Name As String
            Public Sub New(id As Integer, name As String)
                Me.ID = id
                Me.Name = name
            End Sub
        End Class

        Class ProductIDComparer
            Implements IEqualityComparer(Of Product)
            Public Function Equals(x As Product, y As Product) As Boolean Implements IEqualityComparer(Of Product).Equals
                If x Is y Then Return True
                If x Is Nothing OrElse y Is Nothing Then Return False
                Return x.ID = y.ID
            End Function
            Public Function GetHashCode(obj As Product) As Integer Implements IEqualityComparer(Of Product).GetHashCode
                If obj Is Nothing Then Return 0
                Return obj.ID.GetHashCode()
            End Function
        End Class

        Module Program
            Sub Main()
                Dim prods = {New Product(1, "P1"), New Product(1, "P1_Dup"), New Product(2, "P2")}
                Dim unique = prods.Distinct(New ProductIDComparer())
                For Each p In unique
                    __P(CStr(p.ID & "=" & p.Name))
                Next
                __Check("1=P1
2=P2")
            End Sub
        End Module
