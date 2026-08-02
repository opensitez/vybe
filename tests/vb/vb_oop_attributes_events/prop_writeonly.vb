' vybe-test: vb/vb_oop_attributes_events/prop_writeonly
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

Class C: Private _p As Integer: Public WriteOnly Property P As Integer: Set(v As Integer): _p = v: End Set: End Property: Public Function GetP() As Integer: Return _p: End Function: End Class: Module M: Sub Main(): Dim obj As New C(): obj.P = 20: __Check(CStr(obj.GetP()), "20"): End Sub: End Module
