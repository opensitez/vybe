' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_custom_struct_enumerator_pattern
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System

Structure CustomList
    Private arr As Integer()
    Public Sub New(a As Integer())
        arr = a
    End Sub

    Public Function GetEnumerator() As CustomEnumerator
        Return New CustomEnumerator(arr)
    End Function
End Structure

Structure CustomEnumerator
    Private arr As Integer()
    Private idx As Integer
    Public Sub New(a As Integer())
        arr = a
        idx = -1
    End Sub

    Public Function MoveNext() As Boolean
        idx += 1
        Return idx < arr.Length
    End Function

    Public ReadOnly Property Current As Integer
        Get
            Return arr(idx)
        End Get
    End Property
End Structure

Module Program
    Sub Main()
        Dim list As New CustomList(New Integer() {10, 20, 30})
        Dim sum = 0
        For Each item In list
            sum += item
        Next
        Console.WriteLine(sum)
    End Sub
End Module
