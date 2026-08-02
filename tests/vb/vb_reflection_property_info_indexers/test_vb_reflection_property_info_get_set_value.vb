' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_get_set_value
' origin: languages/vb/tests/vb/test_vb_reflection_property_info_indexers.rs

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

Class Account
    Public Property Owner As String = "Alice"
End Class

Module Program
    Sub Main()
        Dim acc As New Account()
        Dim prop = GetType(Account).GetProperty("Owner")
        __Check(CStr(prop.GetValue(acc)), "Alice")
        prop.SetValue(acc, "Bob")
        __Check(CStr(acc.Owner), "Bob")
    End Sub
End Module
