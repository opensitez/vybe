' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_field_enum_isolation
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

Enum State
    Off = 0
    OnVal = 1
End Enum

Class StateHolder(Of T)
    Public Shared CurrentState As State = State.Off
End Class

Module Program
    Sub Main()
        StateHolder(Of Integer).CurrentState = State.OnVal
        __Check(CStr(StateHolder(Of Integer).CurrentState.ToString() & "|" & StateHolder(Of String).CurrentState.ToString()), "OnVal|Off")
    End Sub
End Module
