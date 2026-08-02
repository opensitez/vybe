' vybe-test: vb/vb_property_access_modifiers/property_access_modifiers
' origin: languages/vb/tests/vb/test_vb_property_access_modifiers.rs

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

Class Counter
    Private _count As Integer
    
    ' Property is public, but Set is private
    Public Property Count As Integer
        Get
            Return _count
        End Get
        Private Set(value As Integer)
            _count = value
        End Set
    End Property
    
    Public Sub Increment()
        Count += 1
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Counter()
        c.Increment()
        __Check(CStr(c.Count), "1")
    End Sub
End Module
