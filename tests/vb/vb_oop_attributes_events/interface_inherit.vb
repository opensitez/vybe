' vybe-test: vb/vb_oop_attributes_events/interface_inherit
' origin: languages/vb/tests/vb/test_vb_oop_attributes_events.rs

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

Interface I1: Sub T1(): End Interface: Interface I2: Inherits I1: Sub T2(): End Interface: Class C: Implements I2: Public Sub T1() Implements I2.T1: End Sub: Public Sub T2() Implements I2.T2: End Sub: End Class: Module M: Sub Main(): __Check(CStr("Parsed"), "Parsed"): End Sub: End Module
