' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notoverridable_multiple_overloads
' origin: languages/vb/tests/vb/test_vb_class_sealed_notinheritable_checks.rs

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

Class BaseCalc
    Public Overridable Function Compute(x As Integer) As Integer
        Return x * 2
    End Function
    Public Overridable Function Compute(x As Double) As Double
        Return x * 2.0
    End Function
End Class

Class SealedCalc
    Inherits BaseCalc
    Public NotOverridable Overrides Function Compute(x As Integer) As Integer
        Return x * 3
    End Function
End Class

Module Program
    Sub Main()
        Dim b As BaseCalc = New SealedCalc()
        __Check(CStr(b.Compute(10) & "|" & b.Compute(10.0)), "30|20")
    End Sub
End Module
