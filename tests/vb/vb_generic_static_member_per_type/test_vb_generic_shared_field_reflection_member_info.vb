' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_field_reflection_member_info
' origin: languages/vb/tests/vb/test_vb_generic_static_member_per_type.rs

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

Class ReflectClass(Of T)
    Public Shared Counter As Integer = 0
End Class

Module Program
    Sub Main()
        Dim tInt = GetType(ReflectClass(Of Integer))
        Dim tStr = GetType(ReflectClass(Of String))
        Dim fieldInt = tInt.GetField("Counter")
        Dim fieldStr = tStr.GetField("Counter")

        fieldInt.SetValue(Nothing, 10)
        fieldStr.SetValue(Nothing, 20)

        __Check(CStr(ReflectClass(Of Integer).Counter & "|" & ReflectClass(Of String).Counter), "10|20")
    End Sub
End Module
