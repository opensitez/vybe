' vybe-test: vb/vb_generic_covariance_contravariance_adv/test_vb_generic_covariance_out_modifier
' origin: languages/vb/tests/vb/test_vb_generic_covariance_contravariance_adv.rs

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

Public Interface IReadOnlyContainer(Of Out T)
    Function GetItem() As T
End Interface

Class StringContainer
    Implements IReadOnlyContainer(Of String)
    Public Function GetItem() As String Implements IReadOnlyContainer(Of String).GetItem
        Return "CovariantString"
    End Function
End Class

Module Program
    Sub Main()
        Dim strCont As IReadOnlyContainer(Of String) = New StringContainer()
        Dim objCont As IReadOnlyContainer(Of Object) = strCont ' Covariance assignment
        __Check(CStr(objCont.GetItem()), "CovariantString")
    End Sub
End Module
