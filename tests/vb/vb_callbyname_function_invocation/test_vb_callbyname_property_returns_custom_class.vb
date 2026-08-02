' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_property_returns_custom_class
' origin: languages/vb/tests/vb/test_vb_callbyname_function_invocation.rs

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

Imports Microsoft.VisualBasic

Class SubSystem
    Public Property Name As String = "Core"
End Class

Class RootSystem
    Public Property SubSys As New SubSystem()
End Class

Module Program
    Sub Main()
        Dim r As New RootSystem()
        Dim subObj = CallByName(r, "SubSys", CallType.Get)
        Dim subName = CallByName(subObj, "Name", CallType.Get)
        __Check(CStr(subName), "Core")
    End Sub
End Module
