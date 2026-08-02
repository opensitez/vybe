' vybe-test: vb/vb_abstract_class_inheritance_chain/test_vb_mustinherit_concrete_base_methods
' origin: languages/vb/tests/vb/test_vb_abstract_class_inheritance_chain.rs

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

MustInherit Class BaseLogger
    Public Sub Log(msg As String)
        WriteEntry(FormatMessage(msg))
    End Sub

    Protected MustOverride Sub WriteEntry(formatted As String)

    Protected Virtual Function FormatMessage(msg As String) As String
        Return "[LOG] " & msg
    End Function
End Class

Class ConsoleLogger
    Inherits BaseLogger
    Protected Overrides Sub WriteEntry(formatted As String)
        __Check(CStr(formatted), "[LOG] System initialized")
    End Sub
End Class

Module Program
    Sub Main()
        Dim logger As BaseLogger = New ConsoleLogger()
        logger.Log("System initialized")
    End Sub
End Module
