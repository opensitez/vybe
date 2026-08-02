' vybe-test: vb/vb_class_mybase/class_mybase_constructor
' origin: languages/vb/tests/vb/test_vb_class_mybase.rs

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

Class BaseObj
    Public ID As Integer
    Public Sub New(id As Integer)
        Me.ID = id
    End Sub
End Class

Class DerivedObj
    Inherits BaseObj
    Public Name As String
    
    Public Sub New(id As Integer, name As String)
        MyBase.New(id)
        Me.Name = name
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New DerivedObj(42, "Test")
        __Check(CStr(d.ID), "42")
        __Check(CStr(d.Name), "Test")
    End Sub
End Module
