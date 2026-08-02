' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_event_per_type_isolation
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

Imports System

Class EventBus(Of T)
    Public Shared Event OnEvent As Action(Of T)
    Public Shared Sub Fire(item As T)
        RaiseEvent OnEvent(item)
    End Sub
End Class

Module Program
    Sub Main()
        AddHandler EventBus(Of Integer).OnEvent, Sub(i) __Check(CStr("IntBus: " & i), "IntBus: 42")
        AddHandler EventBus(Of String).OnEvent, Sub(s) __Check(CStr("StringBus: " & s), "StringBus: Hello")

        EventBus(Of Integer).Fire(42)
        EventBus(Of String).Fire("Hello")
    End Sub
End Module
