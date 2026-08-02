' vybe-test: vb/vb_module_properties/module_properties
' origin: languages/vb/tests/vb/test_vb_module_properties.rs

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

Module GlobalState
    Private _val As Integer = 10
    
    ' Properties in a module are implicitly Shared (static)
    Public Property Value As Integer
        Get
            Return _val
        End Get
        Set(val As Integer)
            _val = val
        End Set
    End Property
End Module

Module M
    Sub Main()
        __Check(CStr(GlobalState.Value), "10")
        GlobalState.Value = 20
        __Check(CStr(GlobalState.Value), "20")
    End Sub
End Module
