' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_method_info_create_delegate
' origin: languages/vb/tests/vb/test_vb_reflection_method_info_generic_invoke.rs

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

Class ActionRunner
    Public Function Execute(msg As String) As String
        Return "Executed: " & msg
    End Function
End Class

Module Program
    Sub Main()
        Dim runner As New ActionRunner()
        Dim m = GetType(ActionRunner).GetMethod("Execute")
        Dim del = CType(m.CreateDelegate(GetType(Func(Of String, String)), runner), Func(Of String, String))
        __Check(CStr(del("DirectDelegate")), "Executed: DirectDelegate")
    End Sub
End Module
