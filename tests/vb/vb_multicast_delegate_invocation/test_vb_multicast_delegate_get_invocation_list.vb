' vybe-test: vb/vb_multicast_delegate_invocation/test_vb_multicast_delegate_get_invocation_list
' origin: languages/vb/tests/vb/test_vb_multicast_delegate_invocation.rs

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

Public Delegate Sub SimpleAction()

Module Program
    Sub HandlerA()
    End Sub
    Sub HandlerB()
    End Sub

    Sub Main()
        Dim d As SimpleAction = AddressOf HandlerA
        d = CType([Delegate].Combine(d, AddressOf HandlerB), SimpleAction)
        Dim list As [Delegate]() = d.GetInvocationList()
        __Check(CStr(list.Length), "2")
    End Sub
End Module
