' vybe-test: vb/vb_static_in_generic/static_in_generic
' origin: languages/vb/tests/vb/test_vb_static_in_generic.rs

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
    Function GetCounter(Of T)() As Integer
        ' Static variables inside generic methods are scoped per generic type parameter
        Static c As Integer = 0
        c += 1
        Return c
    End Function

    Sub Main()
        __Check(CStr(GetCounter(Of Integer)()), "1")
        __Check(CStr(GetCounter(Of Integer)()), "2")
        __Check(CStr(GetCounter(Of String)()), "1")
    End Sub
End Module
