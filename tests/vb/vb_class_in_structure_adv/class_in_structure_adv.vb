' vybe-test: vb/vb_class_in_structure_adv/class_in_structure_adv
' origin: languages/vb/tests/vb/test_vb_class_in_structure_adv.rs

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

Structure Outer
    Public Class Inner
        Public Sub Run()
            __Check(CStr("InnerClass"), "InnerClass")
        End Sub
    End Class
End Structure

Module M
    Sub Main()
        Dim i As New Outer.Inner()
        i.Run()
    End Sub
End Module
