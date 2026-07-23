use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Iterator Functions, Yield Return & Yield Exit/Break
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_yield_return_sequence_generator() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function GenerateNumbers() As IEnumerable(Of Integer)
        Yield 10
        Yield 20
        Yield 30
    End Function

    Sub Main()
        Dim list As New List(Of Integer)(GenerateNumbers())
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_yield_return_inside_for_loop() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function RangeGenerator(startVal As Integer, count As Integer) As IEnumerable(Of Integer)
        For i As Integer = 0 To count - 1
            Yield startVal + i
        Next
    End Function

    Sub Main()
        Dim items = RangeGenerator(100, 4)
        Console.WriteLine(String.Join("-", items))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100-101-102-103"]);
}

#[test]
fn test_vb_yield_break_early_exit() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function GeneratorWithBreak() As IEnumerable(Of Integer)
        Yield 1
        Yield 2
        Return ' Return in iterator function acts as Exit Function / yield break!
        Yield 3
    End Function

    Sub Main()
        Dim list As New List(Of Integer)(GeneratorWithBreak())
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2"]);
}

#[test]
fn test_vb_yield_exit_function_early_termination() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function GeneratorWithExit() As IEnumerable(Of String)
        Yield "A"
        If True Then Exit Function
        Yield "B"
    End Function

    Sub Main()
        Dim items = GeneratorWithExit()
        Console.WriteLine(String.Join("", items))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A"]);
}

#[test]
fn test_vb_iterator_function_state_preservation() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function FibonacciSequence(limit As Integer) As IEnumerable(Of Integer)
        Dim a = 0
        Dim b = 1
        For i As Integer = 1 To limit
            Yield a
            Dim temp = a + b
            a = b
            b = temp
        Next
    End Function

    Sub Main()
        Dim fibs = FibonacciSequence(6)
        Console.WriteLine(String.Join(",", fibs))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,1,1,2,3,5"]);
}

#[test]
fn test_vb_iterator_function_returning_ienumerator() {
    let src = r#"
Imports System.Collections

Module Program
    Private Iterator Function GetEnumeratorDirect() As IEnumerator
        Yield "First"
        Yield "Second"
    End Function

    Sub Main()
        Dim enumr = GetEnumeratorDirect()
        While enumr.MoveNext()
            Console.WriteLine(enumr.Current.ToString())
        End While
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First", "Second"]);
}

#[test]
fn test_vb_iterator_with_try_finally_cleanup() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function GeneratorWithFinally() As IEnumerable(Of Integer)
        Try
            Yield 100
            Yield 200
        Finally
            Console.WriteLine("Iterator Finally Executed")
        End Try
    End Function

    Sub Main()
        For Each item In GeneratorWithFinally()
            Console.WriteLine("Item: " & item)
        Next
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Item: 100", "Item: 200", "Iterator Finally Executed"]
    );
}

#[test]
fn test_vb_iterator_partial_consumption_runs_finally() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function InfiniteWithFinally() As IEnumerable(Of Integer)
        Try
            Dim i = 1
            While True
                Yield i
                i += 1
            End While
        Finally
            Console.WriteLine("Cleaned Up Infinite Generator")
        End Try
    End Function

    Sub Main()
        For Each num In InfiniteWithFinally()
            Console.WriteLine("Got: " & num)
            If num = 2 Then Exit For
        Next
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Got: 1", "Got: 2", "Cleaned Up Infinite Generator"]
    );
}

#[test]
fn test_vb_yield_return_generic_struct() {
    let src = r#"
Imports System.Collections.Generic

Structure Pair
    Public Key As String
    Public Val As Integer
End Structure

Module Program
    Private Iterator Function GeneratePairs() As IEnumerable(Of Pair)
        Yield New Pair With {.Key = "K1", .Val = 10}
        Yield New Pair With {.Key = "K2", .Val = 20}
    End Function

    Sub Main()
        For Each p In GeneratePairs()
            Console.WriteLine(p.Key & "=" & p.Val)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["K1=10", "K2=20"]);
}

#[test]
fn test_vb_iterator_lambda_expression() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        ' Iterator lambda expression syntax
        Dim gen As Func(Of IEnumerable(Of Integer)) = Iterator Function()
            Yield 5
            Yield 10
        End Function

        Console.WriteLine(String.Join("+", gen()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5+10"]);
}

#[test]
fn test_vb_iterator_method_in_class_instance() {
    let src = r#"
Imports System.Collections.Generic

Class DataPipeline
    Private data As String() = {"X", "Y", "Z"}

    Public Iterator Function GetFilteredData() As IEnumerable(Of String)
        For Each d In data
            If d <> "Y" Then Yield d
        Next
    End Function
End Class

Module Program
    Sub Main()
        Dim pipe As New DataPipeline()
        Console.WriteLine(String.Join("", pipe.GetFilteredData()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["XZ"]);
}

#[test]
fn test_vb_iterator_recursive_flattening() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function FlattenTree(nodeValue As Integer, depth As Integer) As IEnumerable(Of Integer)
        Yield nodeValue
        If depth > 0 Then
            For Each child In FlattenTree(nodeValue * 10, depth - 1)
                Yield child
            Next
        End If
    End Function

    Sub Main()
        Console.WriteLine(String.Join(",", FlattenTree(1, 2)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,10,100"]);
}

#[test]
fn test_vb_iterator_empty_generator() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function EmptyGen() As IEnumerable(Of String)
        If False Then Yield "Never"
    End Function

    Sub Main()
        Dim list As New List(Of String)(EmptyGen())
        Console.WriteLine(list.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_iterator_with_multiple_exit_points() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function MultiExit(mode As Integer) As IEnumerable(Of String)
        Yield "Step1"
        If mode = 1 Then Return
        Yield "Step2"
        If mode = 2 Then Exit Function
        Yield "Step3"
    End Function

    Sub Main()
        Console.WriteLine(String.Join("|", MultiExit(1)))
        Console.WriteLine(String.Join("|", MultiExit(2)))
        Console.WriteLine(String.Join("|", MultiExit(3)))
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Step1", "Step1|Step2", "Step1|Step2|Step3"]
    );
}

#[test]
fn test_vb_iterator_multiple_enumerators_independent_state() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function CounterGen() As IEnumerable(Of Integer)
        Yield 1
        Yield 2
    End Function

    Sub Main()
        Dim enumerable = CounterGen()
        Dim e1 = enumerable.GetEnumerator()
        Dim e2 = enumerable.GetEnumerator()

        e1.MoveNext()
        Console.WriteLine(e1.Current & "|" & e2.MoveNext() & "|" & e2.Current)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|True|1"]);
}

#[test]
fn test_vb_iterator_throwing_exception_propagates_on_movenext() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Private Iterator Function FaultyGen() As IEnumerable(Of Integer)
        Yield 1
        Throw New InvalidOperationException("Iterator Fault")
    End Function

    Sub Main()
        Try
            For Each item In FaultyGen()
                Console.WriteLine("Item: " & item)
            Next
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Item: 1", "Iterator Fault"]);
}

#[test]
fn test_vb_yield_return_inside_select_case() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function SelectGen(val As Integer) As IEnumerable(Of String)
        Select Case val
            Case 1
                Yield "One"
            Case 2
                Yield "TwoA"
                Yield "TwoB"
            Case Else
                Yield "Other"
        End Select
    End Function

    Sub Main()
        Console.WriteLine(String.Join(",", SelectGen(2)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TwoA,TwoB"]);
}

#[test]
fn test_vb_iterator_property_getter() {
    let src = r#"
Imports System.Collections.Generic

Class SequenceProvider
    Public ReadOnly Iterator Property Sequence As IEnumerable(Of Integer)
        Get
            Yield 100
            Yield 200
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim p As New SequenceProvider()
        Console.WriteLine(String.Join("+", p.Sequence))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100+200"]);
}

#[test]
fn test_vb_iterator_linq_interop_chaining() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Private Iterator Function Numbers() As IEnumerable(Of Integer)
        For i As Integer = 1 To 10
            Yield i
        Next
    End Function

    Sub Main()
        Dim evens = Numbers().Where(Function(n) n Mod 2 = 0).Take(3)
        Console.WriteLine(String.Join(",", evens))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,4,6"]);
}

#[test]
fn test_vb_yield_return_nested_loops() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Iterator Function GridPoints() As IEnumerable(Of String)
        For r As Integer = 0 To 1
            For c As Integer = 0 To 1
                Yield r & ":" & c
            Next
        Next
    End Function

    Sub Main()
        Console.WriteLine(String.Join(" ", GridPoints()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0:0 0:1 1:0 1:1"]);
}
