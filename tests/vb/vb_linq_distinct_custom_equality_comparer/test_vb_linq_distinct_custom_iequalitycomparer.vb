' vybe-test: vb/vb_linq_distinct_custom_equality_comparer/test_vb_linq_distinct_custom_iequalitycomparer
' origin: languages/vb/tests/vb/test_vb_linq_distinct_custom_equality_comparer.rs

Imports System.Collections.Generic
Imports System.Linq

Class Product
    Public Property ID As Integer
    Public Property Name As String
    Public Sub New(id As Integer, name As String) : Me.ID = id : Me.Name = name : End Sub
End Class

Class ProductIDComparer
    Implements IEqualityComparer(Of Product)
    Public Function Equals(x As Product, y As Product) As Boolean Implements IEqualityComparer(Of Product).Equals
        If x Is y Then Return True
        If x Is Nothing OrElse y Is Nothing Then Return False
        Return x.ID = y.ID
    End Function
    Public Function GetHashCode(obj As Product) As Integer Implements IEqualityComparer(Of Product).GetHashCode
        If obj Is Nothing Then Return 0
        Return obj.ID.GetHashCode()
    End Function
End Class

Module Program
    Sub Main()
        Dim prods = {New Product(1, "P1"), New Product(1, "P1_Dup"), New Product(2, "P2")}
        Dim unique = prods.Distinct(New ProductIDComparer())
        For Each p In unique
            Console.WriteLine(p.ID & "=" & p.Name)
        Next
    End Sub
End Module
