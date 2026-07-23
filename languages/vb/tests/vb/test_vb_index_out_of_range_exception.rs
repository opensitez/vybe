use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: IndexOutOfRangeException & Bounds Checks Surface
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_negative_index_throws_index_out_of_range() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Try
            Dim x As Integer = arr(-1)
            Console.WriteLine(x)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Negative Index Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Negative Index Caught"]);
}

#[test]
fn test_vb_array_exceed_upper_bound_throws_index_out_of_range() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Try
            Dim x As Integer = arr(3)
            Console.WriteLine(x)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Exceeded Upper Bound Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Exceeded Upper Bound Caught"]);
}

#[test]
fn test_vb_multidimensional_array_dimension_0_out_of_bounds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim mat(1, 1) As Integer
        Try
            mat(2, 0) = 5
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Dim0 Out Of Bounds Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Dim0 Out Of Bounds Caught"]);
}

#[test]
fn test_vb_multidimensional_array_dimension_1_out_of_bounds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim mat(1, 1) As Integer
        Try
            mat(0, 5) = 10
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Dim1 Out Of Bounds Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Dim1 Out Of Bounds Caught"]);
}

#[test]
fn test_vb_string_char_index_out_of_bounds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim text As String = "Hello"
        Try
            Dim ch As Char = text(10)
            Console.WriteLine(ch)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("String Index Out Of Range Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["String Index Out Of Range Caught"]);
}

#[test]
fn test_vb_empty_array_access_throws_index_out_of_range() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {}
        Try
            Dim x As Integer = arr(0)
            Console.WriteLine(x)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Empty Array Access Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Empty Array Access Caught"]);
}

#[test]
fn test_vb_for_loop_fencepost_error_prevention() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = {"A", "B", "C"}
        ' Valid 0 To arr.Length - 1
        For i As Integer = 0 To arr.Length - 1
            Console.WriteLine(arr(i))
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A", "B", "C"]);
}

#[test]
fn test_vb_for_loop_upper_bound_syntax() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = {"X", "Y"}
        ' In VB, UBound(arr) equals arr.Length - 1
        For i As Integer = 0 To UBound(arr)
            Console.WriteLine(arr(i))
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X", "Y"]);
}

#[test]
fn test_vb_list_out_of_bounds_throws_argument_out_of_range() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2}
        Try
            Dim x As Integer = list(5)
            Console.WriteLine(x)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("List ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["List ArgumentOutOfRangeException Caught"]);
}

#[test]
fn test_vb_array_get_value_out_of_bounds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Array = New String() {"Alpha", "Beta"}
        Try
            Dim val As Object = arr.GetValue(10)
            Console.WriteLine(val)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("GetValue IndexOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["GetValue IndexOutOfRangeException Caught"]
    );
}

#[test]
fn test_vb_array_set_value_out_of_bounds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Array = New String() {"Alpha", "Beta"}
        Try
            arr.SetValue("Gamma", 99)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("SetValue IndexOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["SetValue IndexOutOfRangeException Caught"]
    );
}

#[test]
fn test_vb_array_copy_out_of_bounds_throws_argument_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim srcArr As Integer() = {1, 2, 3, 4, 5}
        Dim destArr As Integer() = new Integer(2) {}
        Try
            ' Trying to copy 5 elements into destination of size 3
            Array.Copy(srcArr, destArr, 5)
        Catch ex As ArgumentException
            Console.WriteLine("Array.Copy ArgumentException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Array.Copy ArgumentException Caught"]);
}

#[test]
fn test_vb_string_substring_index_out_of_bounds_throws_argument_out_of_range() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim str As String = "ABC"
        Try
            Dim subStr As String = str.Substring(0, 10)
            Console.WriteLine(subStr)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("Substring ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Substring ArgumentOutOfRangeException Caught"]
    );
}

#[test]
fn test_vb_indexed_property_out_of_bounds_custom_handling() {
    let src = r#"
Imports System

Class SafeArray
    Private data(2) As Integer
    Default Public Property Item(idx As Integer) As Integer
        Get
            If idx < 0 OrElse idx >= data.Length Then
                Throw New IndexOutOfRangeException("SafeArray index out of bounds")
            End If
            Return data(idx)
        End Get
        Set(value As Integer)
            If idx < 0 OrElse idx >= data.Length Then
                Throw New IndexOutOfRangeException("SafeArray index out of bounds")
            End If
            data(idx) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim sa As New SafeArray()
        Try
            sa(10) = 42
        Catch ex As IndexOutOfRangeException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SafeArray index out of bounds"]);
}

#[test]
fn test_vb_jagged_array_null_sub_array_access() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim jagged As Integer()() = New Integer(2)() {}
        ' Sub-array at index 0 is Nothing
        Try
            Dim val As Integer = jagged(0)(0)
            Console.WriteLine(val)
        Catch ex As NullReferenceException
            Console.WriteLine("Null Sub-Array NullReferenceException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Null Sub-Array NullReferenceException Caught"]
    );
}

#[test]
fn test_vb_jagged_array_sub_array_index_out_of_bounds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim jagged As Integer()() = New Integer(1)() {}
        jagged(0) = New Integer() {10, 20}
        Try
            Dim val As Integer = jagged(0)(5)
            Console.WriteLine(val)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Jagged Sub-Array IndexOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Jagged Sub-Array IndexOutOfRangeException Caught"]
    );
}

#[test]
fn test_vb_read_only_collection_bounds_check() {
    let src = r#"
Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"Item1"}
        Dim ro As ReadOnlyCollection(Of String) = list.AsReadOnly()
        Try
            Dim val = ro(2)
            Console.WriteLine(val)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ReadOnlyCollection ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ReadOnlyCollection ArgumentOutOfRangeException Caught"]
    );
}

#[test]
fn test_vb_span_slice_bounds_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Dim span As Span(Of Integer) = arr.AsSpan()
        Try
            Dim subSpan = span.Slice(1, 5)
            Console.WriteLine(subSpan.Length)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("Span.Slice ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Span.Slice ArgumentOutOfRangeException Caught"]
    );
}

#[test]
fn test_vb_match_collection_regex_index_out_of_bounds() {
    let src = r#"
Imports System
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim matches = Regex.Matches("123 456", "\d+")
        Try
            Dim m = matches(10)
            Console.WriteLine(m.Value)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("MatchCollection ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["MatchCollection ArgumentOutOfRangeException Caught"]
    );
}

#[test]
fn test_vb_safe_array_element_getter_extension() {
    let src = r#"
Imports System.Runtime.CompilerServices

Module ArrayExtensions
    <Extension()>
    Public Function ElementAtOrDefault(Of T)(arr As T(), index As Integer, defaultValue As T) As T
        If arr IsNot Nothing AndAlso index >= 0 AndAlso index < arr.Length Then
            Return arr(index)
        End If
        Return defaultValue
    End Function
End Module

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Console.WriteLine(arr.ElementAtOrDefault(1, -1) & "|" & arr.ElementAtOrDefault(5, -1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20|-1"]);
}
