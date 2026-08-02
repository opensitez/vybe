' vybe-test: vb/vb_generics_classes/generic_class_basic
' origin: languages/vb/tests/vb/test_vb_generics_classes.rs

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

Class Box(Of T)
    Private _value As T
    
    Public Sub New(val As T)
        _value = val
    End Sub
    
    Public Function GetValue() As T
        Return _value
    End Function
End Class

Module M
    Sub Main()
        Dim intBox As New Box(Of Integer)(42)
        Dim strBox As New Box(Of String)("Hello")
        
        __Check(CStr(intBox.GetValue()), "42")
        __Check(CStr(strBox.GetValue()), "Hello")
    End Sub
End Module
