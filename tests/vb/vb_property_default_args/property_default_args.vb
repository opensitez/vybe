' vybe-test: vb/vb_property_default_args/property_default_args
' origin: languages/vb/tests/vb/test_vb_property_default_args.rs

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

Class Cache
    Private _items As New System.Collections.Generic.Dictionary(Of String, String)
    
    ' Default Property allows the object to be indexed directly like an array
    Default Public Property Item(key As String) As String
        Get
            If _items.ContainsKey(key) Then Return _items(key)
            Return Nothing
        End Get
        Set(value As String)
            _items(key) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache()
        ' Using the default property
        c("A") = "Apple"
        c("B") = "Banana"
        
        __Check(CStr(c("A")), "Apple")
        __Check(CStr(c("B")), "Banana")
    End Sub
End Module
