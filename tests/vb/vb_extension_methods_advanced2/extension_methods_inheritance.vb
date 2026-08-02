' vybe-test: vb/vb_extension_methods_advanced2/extension_methods_inheritance
' origin: languages/vb/tests/vb/test_vb_extension_methods_advanced2.rs

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

Imports System.Runtime.CompilerServices

Interface IEntity
    ReadOnly Property Id As Integer
End Interface

Class User
    Implements IEntity
    Public ReadOnly Property Id As Integer Implements IEntity.Id
        Get
            Return 42
        End Get
    End Property
End Class

Module ExtensionMethods
    <Extension()>
    Public Function GetIdentifier(entity As IEntity) As String
        Return "Entity-" & entity.Id.ToString()
    End Function
End Module

Module M
    Sub Main()
        Dim u As New User()
        ' Extension method on interface
        __Check(CStr(u.GetIdentifier()), "Entity-42")
    End Sub
End Module
