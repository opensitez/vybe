' vybe-test: vb/vb_attributes_class/attribute_class_basic
' origin: languages/vb/tests/vb/test_vb_attributes_class.rs

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

<Serializable>
Class DataHolder
    Public Value As String = "Test"
End Class

Module M
    Sub Main()
        Dim t As Type = GetType(DataHolder)
        ' Check if attribute is applied
        Dim isSerializable As Boolean = t.IsSerializable
        __Check(CStr(isSerializable), "True")
    End Sub
End Module
