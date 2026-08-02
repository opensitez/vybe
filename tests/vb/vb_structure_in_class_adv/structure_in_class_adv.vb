' vybe-test: vb/vb_structure_in_class_adv/structure_in_class_adv
' origin: languages/vb/tests/vb/test_vb_structure_in_class_adv.rs

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

Class Outer
    Public Structure Inner
        Public Val As Integer
    End Structure
End Class

Module M
    Sub Main()
        Dim i As New Outer.Inner()
        i.Val = 10
        __Check(CStr(i.Val), "10")
    End Sub
End Module
