' vybe-test: vb/vb_addressof_interface_method/addressof_interface_method
' origin: languages/vb/tests/vb/test_vb_addressof_interface_method.rs

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

Interface IWorker
    Sub Work()
End Interface

Class Worker
    Implements IWorker
    
    Public Sub Work() Implements IWorker.Work
        __Check(CStr("InterfaceWork"), "InterfaceWork")
    End Sub
End Class

Module M
    Sub Main()
        Dim w As IWorker = New Worker()
        
        ' AddressOf through an interface
        Dim act As Action = AddressOf w.Work
        act()
    End Sub
End Module
