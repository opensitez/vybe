' vybe-test: vb/vb_oop_attributes_events/attr_custom
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

<AttributeUsage(AttributeTargets.Class)> Class MyAttr: Inherits Attribute: Public Name As String: Public Sub New(n As String): Name = n: End Sub: End Class: <MyAttr("Test")> Class C: End Class: Module M: Sub Main(): Dim a = CType(GetType(C).GetCustomAttributes(False)(0), MyAttr): __Check(CStr(a.Name), "Test"): End Sub: End Module
