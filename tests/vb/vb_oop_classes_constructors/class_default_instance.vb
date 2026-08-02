' vybe-test: vb/vb_oop_classes_constructors/class_default_instance
' origin: languages/vb/tests/vb/test_vb_oop_classes_constructors.rs

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

Module M
Sub Main()
' VB has default instances for Forms (My.Forms.Form1), but for normal classes it requires special attributes
__Check(CStr("Parsed"), "Parsed")
End Sub
End Module
