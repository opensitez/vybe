' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_get_invocation_list_via_reflection
' origin: languages/vb/tests/vb/test_vb_event_raise_multithread_safe.rs

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
Imports System.Reflection

Class Subject
    Public Event Update As EventHandler
    Public Function GetSubscriberCount() As Integer
        Dim field As FieldInfo = GetType(Subject).GetField("UpdateEvent", BindingFlags.NonPublic Or BindingFlags.Instance)
        If field IsNot Nothing Then
            Dim del = TryCast(field.GetValue(Me), [Delegate])
            If del IsNot Nothing Then Return del.GetInvocationList().Length
        End If
        Return 0
    End Function
End Class

Module Program
    Sub Main()
        Dim s As New Subject()
        AddHandler s.Update, Sub(sender, args) End Sub
        AddHandler s.Update, Sub(sender, args) End Sub
        __Check(CStr(s.GetSubscriberCount()), "2")
    End Sub
End Module
