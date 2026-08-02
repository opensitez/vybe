' vybe-test: vb/vb_multicast_delegate_invocation/test_vb_multicast_delegate_combine_remove
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

Public Delegate Sub NotifyHandler(msg As String)

Class Notifier
    Public Shared Log As String = ""
    Public Shared Sub Method1(m As String)
        Log &= "M1:" & m & ";"
    End Sub
    Public Shared Sub Method2(m As String)
        Log &= "M2:" & m & ";"
    End Sub
End Class

Module Program
    Sub Main()
        Dim d1 As NotifyHandler = AddressOf Notifier.Method1
        Dim d2 As NotifyHandler = AddressOf Notifier.Method2
        Dim multi As NotifyHandler = CType([Delegate].Combine(d1, d2), NotifyHandler)
        multi("Hello")

        Dim singleDel As NotifyHandler = CType([Delegate].Remove(multi, d1), NotifyHandler)
        singleDel("World")

        __Check(CStr(Notifier.Log), "M1:Hello;M2:Hello;M2:World;")
    End Sub
End Module
