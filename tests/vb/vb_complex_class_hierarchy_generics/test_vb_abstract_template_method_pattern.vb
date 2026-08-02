' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_abstract_template_method_pattern
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

MustInherit Class DataProcessor
    Public Sub Process()
        Step1()
        Step2()
    End Sub

    Protected MustOverride Sub Step1()
        Protected MustOverride Sub Step2()
End Class

Class XmlProcessor
    Inherits DataProcessor
    Protected Overrides Sub Step1()
        __Check(CStr("Xml Step 1"), "Xml Step 1")
    End Sub
    Protected Overrides Sub Step2()
        __Check(CStr("Xml Step 2"), "Xml Step 2")
    End Sub
End Class

Module Program
    Sub Main()
        Dim proc As DataProcessor = New XmlProcessor()
        proc.Process()
    End Sub
End Module
