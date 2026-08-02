' vybe-test: vb/vb_class_shadows/class_shadows_method
' origin: languages/vb/tests/vb/test_vb_class_shadows.rs

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

Class Parent
    Public Sub ShowMessage()
        __Check(CStr("Parent Message"), "Child Message")
    End Sub
End Class

Class Child
    Inherits Parent
    
    Public Shadows Sub ShowMessage()
        __Check(CStr("Child Message"), "Parent Message")
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Child()
        c.ShowMessage()
        
        Dim p As Parent = c
        p.ShowMessage() ' Should print Parent Message due to shadowing (not overriding)
    End Sub
End Module
