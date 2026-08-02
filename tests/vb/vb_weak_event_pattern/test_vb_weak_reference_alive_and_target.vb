' vybe-test: vb/vb_weak_event_pattern/test_vb_weak_reference_alive_and_target
' origin: languages/vb/tests/vb/test_vb_weak_event_pattern.rs

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

Class TargetData
    Public Property Value As String = "Alive"
End Class

Module Program
    Sub Main()
        Dim obj As New TargetData()
        Dim weak As New WeakReference(obj)
        __Check(CStr(weak.IsAlive), "True")
        Dim retrieved As TargetData = CType(weak.Target, TargetData)
        __Check(CStr(retrieved.Value), "Alive")
    End Sub
End Module
