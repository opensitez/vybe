use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: For Each Loop & Struct Custom Enumerators Pattern
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_for_each_custom_struct_enumerator_pattern() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["60"]);
}

#[test]
fn test_vb_for_each_array_implicit_type_inference() {
    let src = r#"
Module Program
    Sub Main()
        Dim items As String() = {"Alpha", "Beta", "Gamma"}
        Dim concat = ""
        For Each item In items
            concat &= item
        Next
        Console.WriteLine(concat)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AlphaBetaGamma"]);
}

#[test]
fn test_vb_for_each_multidimensional_array_row_major_order() {
    let src = r#"
Module Program
    Sub Main()
        Dim grid(,) As Integer = {{1, 2}, {3, 4}}
        Dim res = ""
        For Each val In grid
            res &= val & ","
        Next
        Console.WriteLine(res.TrimEnd(","c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,4"]);
}

#[test]
fn test_vb_for_each_generic_ienumerable_implementation() {
    let src = r#"
Imports System.Collections
Imports System.Collections.Generic

Class NumberCollection
    Implements IEnumerable(Of Integer)

    Public Function GetEnumerator() As IEnumerator(Of Integer) Implements IEnumerable(Of Integer).GetEnumerator
        Return New List(Of Integer) From {1, 3, 5}.GetEnumerator()
    End Function

    Private Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function
End Class

Module Program
    Sub Main()
        Dim col As New NumberCollection()
        Dim total = 0
        For Each n In col
            total += n
        Next
        Console.WriteLine(total)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9"]);
}

#[test]
fn test_vb_for_each_with_exit_for() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5}
        Dim lastSeen = 0
        For Each n In numbers
            If n = 4 Then Exit For
            lastSeen = n
        Next
        Console.WriteLine(lastSeen)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_for_each_with_continue_for() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5}
        Dim oddSum = 0
        For Each n In numbers
            If n Mod 2 = 0 Then Continue For
            oddSum += n
        Next
        Console.WriteLine(oddSum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9"]);
}

#[test]
fn test_vb_for_each_dictionary_key_value_pair() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer)()
        dict("A") = 1
        dict("B") = 2

        For Each kvp As KeyValuePair(Of String, Integer) In dict
            Console.WriteLine(kvp.Key & ":" & kvp.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A:1", "B:2"]);
}

#[test]
fn test_vb_for_each_string_characters() {
    let src = r#"
Module Program
    Sub Main()
        Dim text = "Vybe"
        For Each ch As Char In text
            Console.WriteLine(ch)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["V", "y", "b", "e"]);
}

#[test]
fn test_vb_for_each_disposes_disposable_enumerator() {
    let src = r#"
Imports System
Imports System.Collections
Imports System.Collections.Generic

Class DisposableCollection
    Implements IEnumerable(Of String)

    Private Class CustomDispEnum
        Implements IEnumerator(Of String)
        Public Property Current As String Implements IEnumerator(Of String).Current
        Private Property Current1 As Object Implements IEnumerator.Current
            Get
                Return Current
            End Get
        End Property

        Private readDone As Boolean = False
        Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
            If Not readDone Then
                Current = "SingleItem"
                readDone = True
                Return True
            End If
            Return False
        End Function

        Public Sub Reset() Implements IEnumerator.Reset
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Console.WriteLine("Enumerator Disposed")
        End Sub
    End Class

    Public Function GetEnumerator() As IEnumerator(Of String) Implements IEnumerable(Of String).GetEnumerator
        Return New CustomDispEnum()
    End Function

    Private Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function
End Class

Module Program
    Sub Main()
        Dim col As New DisposableCollection()
        For Each item In col
            Console.WriteLine(item)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SingleItem", "Enumerator Disposed"]);
}

#[test]
fn test_vb_for_each_loop_variable_scoping_in_lambda_capture() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim actions As New List(Of Action)()
        For Each item In New String() {"A", "B", "C"}
            actions.Add(Sub() Console.WriteLine(item))
        Next

        For Each act In actions
            act()
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A", "B", "C"]);
}

#[test]
fn test_vb_for_each_empty_collection_does_not_execute_body() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer)()
        Dim executed = False
        For Each item In list
            executed = True
        Next
        Console.WriteLine(executed)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_for_each_nested_loops() {
    let src = r#"
Module Program
    Sub Main()
        Dim letters As String() = {"X", "Y"}
        Dim numbers As Integer() = {1, 2}
        For Each l In letters
            For Each n In numbers
                Console.WriteLine(l & n)
            Next
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X1", "X2", "Y1", "Y2"]);
}

#[test]
fn test_vb_for_each_loop_variable_explicit_type_widening() {
    let src = r#"
Module Program
    Sub Main()
        Dim bytes As Byte() = {1, 2, 3}
        ' Explicitly typed loop variable Double widens from Byte!
        For Each val As Double In bytes
            Console.WriteLine(val.ToString("F1"))
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.0", "2.0", "3.0"]);
}

#[test]
fn test_vb_for_each_custom_duck_typed_collection() {
    let src = r#"
Class DuckCollection
    Public Function GetEnumerator() As DuckEnumerator
        Return New DuckEnumerator()
    End Function
End Class

Class DuckEnumerator
    Private count As Integer = 0
    Public Function MoveNext() As Boolean
        count += 1
        Return count <= 2
    End Function
    Public ReadOnly Property Current As String
        Get
            Return "Quack" & count
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim duckCol As New DuckCollection()
        For Each item In duckCol
            Console.WriteLine(item)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Quack1", "Quack2"]);
}

#[test]
fn test_vb_for_each_exception_during_enumeration() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Private Iterator Function FaultyEnum() As IEnumerable(Of Integer)
        Yield 1
        Throw New InvalidOperationException("Faulty Enum Error")
    End Function

    Sub Main()
        Try
            For Each item In FaultyEnum()
                Console.WriteLine(item)
            Next
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "Faulty Enum Error"]);
}

#[test]
fn test_vb_for_each_over_null_collection_throws_null_reference() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As List(Of Integer) = Nothing
        Try
            For Each item In list
                Console.WriteLine(item)
            Next
        Catch ex As NullReferenceException
            Console.WriteLine("NullReferenceException Caught on For Each Null")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["NullReferenceException Caught on For Each Null"]
    );
}

#[test]
fn test_vb_for_each_array_segment() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim raw As Integer() = {10, 20, 30, 40, 50}
        Dim seg As New ArraySegment(Of Integer)(raw, 1, 3)
        Dim sum = 0
        For Each val In seg
            sum += val
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["90"]);
}

#[test]
fn test_vb_for_each_linked_list_traversal() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New LinkedList(Of String)()
        list.AddLast("Node1")
        list.AddLast("Node2")
        For Each node In list
            Console.WriteLine(node)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Node1", "Node2"]);
}

#[test]
fn test_vb_for_each_hashset_unique_elements() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set As New HashSet(Of Integer) From {1, 2, 2, 3}
        Console.WriteLine(set.Count)
        For Each item In set
            Console.WriteLine(item)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "1", "2", "3"]);
}

#[test]
fn test_vb_for_each_stack_lifo_order() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim s As New Stack(Of String)()
        s.Push("First")
        s.Push("Second")
        For Each item In s
            Console.WriteLine(item)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Second", "First"]);
}
