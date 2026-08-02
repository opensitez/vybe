' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_delegates_as_interface_properties
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface ICallbackContainer
    Property Handler As Action(Of String)
End Interface

Class Worker
    Implements ICallbackContainer
    Public Property Handler As Action(Of String) Implements ICallbackContainer.Handler
    Public Sub Run()
        If Handler IsNot Nothing Then
            Handler("Finished Work")
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Dim w As New Worker()
        Dim c As ICallbackContainer = w
        c.Handler = Sub(msg) __Check(CStr("Callback: " & msg), "Callback: Finished Work")
        w.Run()
    End Sub
End Module
