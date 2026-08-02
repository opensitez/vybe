' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_custom_structure_instance
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

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Function GetSum() As Integer
        Return X + Y
    End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Point With {.X = 10, .Y = 20}
        Dim res = CallByName(p, "GetSum", CallType.Method)
        __Check(CStr(res), "30")
    End Sub
End Module
