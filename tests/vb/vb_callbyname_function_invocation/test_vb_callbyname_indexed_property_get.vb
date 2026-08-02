' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_indexed_property_get
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

Module Program
    Class Catalog
        Private items As String() = {"Alpha", "Beta", "Gamma"}
        Default Public Property Item(idx As Integer) As String
            Get
                Return items(idx)
            End Get
        End Property
    End Class

    Sub Main()
        Dim cat As New Catalog()
        Dim res = CallByName(cat, "Item", CallType.Get, 1)
        __Check(CStr(res), "Beta")
    End Sub
End Module
