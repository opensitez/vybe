' vybe-test: vb/vb_oop_attributes_events/struct_copy
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

Structure S: Public V As Integer: End Structure: Module M: Sub Main(): Dim s1 As New S(): s1.V = 1: Dim s2 = s1: s2.V = 2: __Check(CStr(s1.V), "1"): End Sub: End Module
