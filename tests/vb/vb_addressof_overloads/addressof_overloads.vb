' vybe-test: vb/vb_addressof_overloads/addressof_overloads
' origin: languages/vb/tests/vb/test_vb_addressof_overloads.rs

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

Delegate Sub PrintString(s As String)
Delegate Sub PrintInteger(i As Integer)

Module M
    Sub Print(s As String)
        __Check(CStr("String: " & s), "String: Hello")
    End Sub

    Sub Print(i As Integer)
        __Check(CStr("Integer: " & i.ToString()), "Integer: 42")
    End Sub

    Sub Main()
        ' AddressOf automatically selects the correct overload based on the target delegate type
        Dim ds As PrintString = AddressOf Print
        Dim di As PrintInteger = AddressOf Print
        
        ds("Hello")
        di(42)
    End Sub
End Module
