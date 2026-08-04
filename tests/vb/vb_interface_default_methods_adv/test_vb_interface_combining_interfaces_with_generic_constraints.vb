' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_combining_interfaces_with_generic_constraints
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IValidatable
    Function IsValid() As Boolean
End Interface

Class Processor(Of T As IValidatable)
    Public Function Process(item As T) As String
        If item.IsValid() Then
            Return "Valid"
        Else
            Return "Invalid"
        End If
    End Function
End Class

Class FormInput
    Implements IValidatable
    Public Input As String
    Public Sub New(i As String)
        Input = i
    End Sub
    Public Function IsValid() As Boolean Implements IValidatable.IsValid
        Return Not String.IsNullOrEmpty(Input)
    End Function
End Class

Module Program
    Sub Main()
        Dim p As New Processor(Of FormInput)()
        __P(CStr(p.Process(New FormInput("OK"))))
        __P(CStr(p.Process(New FormInput(""))))
        __Check("Valid
Invalid")
    End Sub
End Module
