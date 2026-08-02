' vybe-test: vb/vb_auto_property_initializers_adv/test_vb_auto_property_readonly_initializer
' origin: languages/vb/tests/vb/test_vb_auto_property_initializers_adv.rs

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

Class ImmutablePoint
    Public ReadOnly Property X As Double = 1.0
    Public ReadOnly Property Y As Double = 2.0
End Class

Module Program
    Sub Main()
        Dim pt As New ImmutablePoint()
        __Check(CStr(pt.X & "," & pt.Y), "1,2")
    End Sub
End Module
