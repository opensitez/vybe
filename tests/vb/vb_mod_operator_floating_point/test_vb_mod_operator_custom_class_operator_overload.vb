' vybe-test: vb/vb_mod_operator_floating_point/test_vb_mod_operator_custom_class_operator_overload
' origin: languages/vb/tests/vb/test_vb_mod_operator_floating_point.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Module Program
    Class ClockTime
        Public Hours As Integer
        Public Sub New(h As Integer)
            Hours = h
        End Sub
        Public Shared Operator Mod(a As ClockTime, b As Integer) As ClockTime
            Return New ClockTime(a.Hours Mod b)
        End Operator
    End Class

    Sub Main()
        Dim t As New ClockTime(27)
        Dim wrapped = t Mod 24
        __Check(CStr(wrapped.Hours), "3")
    End Sub
End Module
