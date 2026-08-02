' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_override_equals_hashcode
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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

Structure Token(Of T)
    Public Data As T
    Public Sub New(d As T)
        Data = d
    End Sub
    Public Overrides Function Equals(obj As Object) As Boolean
        If Not (TypeOf obj Is Token(Of T)) Then Return False
        Dim other = CType(obj, Token(Of T))
        Return Object.Equals(Data, other.Data)
    End Function
    Public Overrides Function GetHashCode() As Integer
        If Data Is Nothing Then Return 0
        Return Data.GetHashCode()
    End Function
End Structure

Module Program
    Sub Main()
        Dim t1 As New Token(Of String)("ABC")
        Dim t2 As New Token(Of String)("ABC")
        Dim t3 As New Token(Of String)("XYZ")
        __Check(CStr(t1.Equals(t2) & "|" & t1.Equals(t3)), "True|False")
    End Sub
End Module
