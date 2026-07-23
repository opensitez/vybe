use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array.Resize, ReDim Preserve & Memory Allocation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_resize_expand_integer_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Array.Resize(arr, 5)
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2) & "," & arr(3) & "," & arr(4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5", "10,20,30,0,0"]);
}

#[test]
fn test_vb_array_resize_shrink_integer_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30, 40, 50}
        Array.Resize(arr, 2)
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(0) & "," & arr(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "10,20"]);
}

#[test]
fn test_vb_array_resize_string_array_defaults_to_nothing() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As String() = {"A", "B"}
        Array.Resize(arr, 4)
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(2) Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4", "True"]);
}

#[test]
fn test_vb_array_redim_preserve_expand_1d() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(2) As Integer
        arr(0) = 1
        arr(1) = 2
        arr(2) = 3

        ReDim Preserve arr(4)
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2) & "," & arr(3) & "," & arr(4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5", "1,2,3,0,0"]);
}

#[test]
fn test_vb_array_redim_preserve_shrink_1d() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(4) As String
        arr(0) = "One"
        arr(1) = "Two"
        arr(2) = "Three"
        arr(3) = "Four"
        arr(4) = "Five"

        ReDim Preserve arr(1)
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(0) & "," & arr(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "One,Two"]);
}

#[test]
fn test_vb_array_redim_without_preserve_clears_data() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(2) As Integer
        arr(0) = 99
        arr(1) = 88
        arr(2) = 77

        ReDim arr(2)
        Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,0,0"]);
}

#[test]
fn test_vb_array_redim_preserve_multidimensional_last_dimension_only() {
    let src = r#"
Module Program
    Sub Main()
        Dim mat(1, 1) As Integer
        mat(0, 0) = 1 : mat(0, 1) = 2
        mat(1, 0) = 3 : mat(1, 1) = 4

        ReDim Preserve mat(1, 2)
        Console.WriteLine(mat(0, 0) & "," & mat(0, 1) & "," & mat(0, 2))
        Console.WriteLine(mat(1, 0) & "," & mat(1, 1) & "," & mat(1, 2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,0", "3,4,0"]);
}

#[test]
fn test_vb_array_resize_null_array_allocates_new() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Double() = Nothing
        Array.Resize(arr, 3)
        Console.WriteLine(arr IsNot Nothing)
        Console.WriteLine(arr.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "3"]);
}

#[test]
fn test_vb_array_resize_to_zero_length() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {1, 2, 3}
        Array.Resize(arr, 0)
        Console.WriteLine(arr.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_array_redim_preserve_boolean_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim flags(1) As Boolean
        flags(0) = True
        ReDim Preserve flags(3)
        Console.WriteLine(flags(0) & "," & flags(1) & "," & flags(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True,False,False"]);
}

#[test]
fn test_vb_array_redim_preserve_struct_array() {
    let src = r#"
Structure Pair
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Sub Main()
        Dim pairs(0) As Pair
        pairs(0).X = 10 : pairs(0).Y = 20
        ReDim Preserve pairs(1)
        Console.WriteLine(pairs(0).X & ":" & pairs(0).Y)
        Console.WriteLine(pairs(1).X & ":" & pairs(1).Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10:20", "0:0"]);
}

#[test]
fn test_vb_array_redim_preserve_reference_type_instances() {
    let src = r#"
Class Item
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

Module Program
    Sub Main()
        Dim items(0) As Item
        items(0) = New Item("First")
        ReDim Preserve items(1)
        Console.WriteLine(items(0).Name)
        Console.WriteLine(items(1) Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First", "True"]);
}

#[test]
fn test_vb_array_redim_preserve_repeated_expansions() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(0) As Integer
        arr(0) = 1
        For i As Integer = 1 To 4
            ReDim Preserve arr(i)
            arr(i) = (i + 1) * 10
        Next
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,20,30,40,50"]);
}

#[test]
fn test_vb_array_resize_generic_list_convert() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr As Byte() = {255, 128}
        Array.Resize(arr, 4)
        Console.WriteLine(String.Join("-", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255-128-0-0"]);
}

#[test]
fn test_vb_array_redim_preserve_char_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim chars(1) As Char
        chars(0) = "A"c
        chars(1) = "B"c
        ReDim Preserve chars(3)
        Console.WriteLine(chars(0) & chars(1) & "|" & (chars(2) = ChrW(0)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AB|True"]);
}

#[test]
fn test_vb_array_redim_preserve_decimal_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim decs(0) As Decimal
        decs(0) = 123.45D
        ReDim Preserve decs(1)
        Console.WriteLine(decs(0) & ":" & decs(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123.45:0"]);
}

#[test]
fn test_vb_array_redim_preserve_date_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dates(0) As DateTime
        dates(0) = New DateTime(2025, 1, 1)
        ReDim Preserve dates(1)
        Console.WriteLine(dates(0).ToString("yyyy-MM-dd"))
        Console.WriteLine(dates(1) = DateTime.MinValue)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-01-01", "True"]);
}

#[test]
fn test_vb_array_redim_preserve_enum_array() {
    let src = r#"
Enum Priority
    Low = 0
    Medium = 1
    High = 2
End Enum

Module Program
    Sub Main()
        Dim priorities(0) As Priority
        priorities(0) = Priority.High
        ReDim Preserve priorities(1)
        Console.WriteLine(priorities(0) & ":" & priorities(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["High:Low"]);
}

#[test]
fn test_vb_array_resize_reference_equality() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim original As Integer() = {1, 2, 3}
        Dim reference As Integer() = original
        Array.Resize(original, 5)
        Console.WriteLine(original.Length)
        Console.WriteLine(reference.Length)
        Console.WriteLine(Object.ReferenceEquals(original, reference))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5", "3", "False"]);
}

#[test]
fn test_vb_array_redim_preserve_object_array_mixed_types() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj(1) As Object
        obj(0) = 42
        obj(1) = "Hello"
        ReDim Preserve obj(2)
        obj(2) = True
        Console.WriteLine(obj(0).ToString() & "|" & obj(1).ToString() & "|" & obj(2).ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42|Hello|True"]);
}
