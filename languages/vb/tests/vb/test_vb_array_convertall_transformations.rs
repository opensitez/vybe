use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array.ConvertAll Element Transformations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_convertall_int_to_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4}
        Dim strings As String() = Array.ConvertAll(numbers, Function(n) "Item_" & n)
        Console.WriteLine(String.Join(",", strings))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Item_1,Item_2,Item_3,Item_4"]);
}

#[test]
fn test_vb_array_convertall_string_to_int() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim strNums As String() = {"10", "20", "30"}
        Dim numbers As Integer() = Array.ConvertAll(strNums, Function(s) Integer.Parse(s))
        Console.WriteLine(numbers(0) + numbers(1) + numbers(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["60"]);
}

#[test]
fn test_vb_array_convertall_double_to_int_truncation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim doubles As Double() = {1.9, 2.1, 3.8}
        Dim ints As Integer() = Array.ConvertAll(doubles, Function(d) CInt(Math.Floor(d)))
        Console.WriteLine(String.Join(",", ints))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_array_convertall_empty_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        Dim converted As String() = Array.ConvertAll(empty, Function(n) n.ToString())
        Console.WriteLine(converted.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_array_convertall_object_instantiation() {
    let src = r#"
Imports System

Class User
    Public ReadOnly Property Name As String
    Public Sub New(name As String)
        Me.Name = name
    End Sub
End Class

Module Program
    Sub Main()
        Dim names As String() = {"Alice", "Bob"}
        Dim users As User() = Array.ConvertAll(names, Function(n) New User(n))
        Console.WriteLine(users(0).Name & "&" & users(1).Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice&Bob"]);
}

#[test]
fn test_vb_array_convertall_extract_property() {
    let src = r#"
Imports System

Class Product
    Public Property Id As Integer
    Public Sub New(id As Integer)
        Me.Id = id
    End Sub
End Class

Module Program
    Sub Main()
        Dim prods As Product() = {New Product(101), New Product(102)}
        Dim ids As Integer() = Array.ConvertAll(prods, Function(p) p.Id)
        Console.WriteLine(String.Join("-", ids))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["101-102"]);
}

#[test]
fn test_vb_array_convertall_boolean_negation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim flags As Boolean() = {True, False, True}
        Dim inverted As Boolean() = Array.ConvertAll(flags, Function(b) Not b)
        Console.WriteLine(String.Join(",", inverted))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False,True,False"]);
}

#[test]
fn test_vb_array_convertall_char_to_ascii_code() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim chars As Char() = {"A"c, "B"c, "C"c}
        Dim ascii As Integer() = Array.ConvertAll(chars, Function(c) AscW(c))
        Console.WriteLine(String.Join(",", ascii))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["65,66,67"]);
}

#[test]
fn test_vb_array_convertall_ascii_code_to_char() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ascii As Integer() = {65, 66, 67}
        Dim chars As Char() = Array.ConvertAll(ascii, Function(i) ChrW(i))
        Console.WriteLine(New String(chars))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ABC"]);
}

#[test]
fn test_vb_array_convertall_enum_to_int() {
    let src = r#"
Imports System

Enum Level
    Low = 1
    Medium = 2
    High = 3
End Enum

Module Program
    Sub Main()
        Dim levels As Level() = {Level.Low, Level.High}
        Dim ints As Integer() = Array.ConvertAll(levels, Function(l) CInt(l))
        Console.WriteLine(String.Join(",", ints))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,3"]);
}

#[test]
fn test_vb_array_convertall_int_to_enum() {
    let src = r#"
Imports System

Enum Mode
    Off = 0
    OnVal = 1
End Enum

Module Program
    Sub Main()
        Dim raw As Integer() = {0, 1, 0}
        Dim modes As Mode() = Array.ConvertAll(raw, Function(r) CType(r, Mode))
        Console.WriteLine(modes(0).ToString() & "," & modes(1).ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Off,OnVal"]);
}

#[test]
fn test_vb_array_convertall_tuple_creation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim keys As String() = {"K1", "K2"}
        Dim tuples As (String, Integer)() = Array.ConvertAll(keys, Function(k) (k, k.Length))
        Console.WriteLine(tuples(0).Item1 & ":" & tuples(0).Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["K1:2"]);
}

#[test]
fn test_vb_array_convertall_struct_transformation() {
    let src = r#"
Imports System

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim coords As Integer() = {10, 20, 30, 40}
        ' Take pairs to points
        Dim pts As Point() = {New Point(10, 20), New Point(30, 40)}
        Dim transformed As Point() = Array.ConvertAll(pts, Function(p) New Point(p.X * 2, p.Y * 2))
        Console.WriteLine(transformed(0).X & "," & transformed(0).Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,40"]);
}

#[test]
fn test_vb_array_convertall_datetime_formatting() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dates As DateTime() = {New DateTime(2025, 1, 1), New DateTime(2025, 12, 31)}
        Dim formatted As String() = Array.ConvertAll(dates, Function(d) d.ToString("yyyy-MM-dd"))
        Console.WriteLine(String.Join(";", formatted))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-01-01;2025-12-31"]);
}

#[test]
fn test_vb_array_convertall_nullable_to_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim nullables As Nullable(Of Integer)() = {10, Nothing, 30}
        Dim values As Integer() = Array.ConvertAll(nullables, Function(n) n.GetValueOrDefault(-1))
        Console.WriteLine(String.Join(",", values))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,-1,30"]);
}

#[test]
fn test_vb_array_convertall_string_to_uppercase() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim words As String() = {"apple", "banana"}
        Dim upper As String() = Array.ConvertAll(words, Function(w) w.ToUpper())
        Console.WriteLine(String.Join(",", upper))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["APPLE,BANANA"]);
}

#[test]
fn test_vb_array_convertall_math_pow_transformation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bases As Double() = {1.0, 2.0, 3.0, 4.0}
        Dim squares As Double() = Array.ConvertAll(bases, Function(b) Math.Pow(b, 2.0))
        Console.WriteLine(String.Join(",", squares))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,4,9,16"]);
}

#[test]
fn test_vb_array_convertall_timespan_seconds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim seconds As Double() = {60.0, 120.0, 180.0}
        Dim timeSpans As TimeSpan() = Array.ConvertAll(seconds, Function(s) TimeSpan.FromSeconds(s))
        Console.WriteLine(timeSpans(0).TotalMinutes & "," & timeSpans(1).TotalMinutes)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2"]);
}

#[test]
fn test_vb_array_convertall_byte_array_to_hex_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = {10, 15, 255}
        Dim hexes As String() = Array.ConvertAll(bytes, Function(b) b.ToString("X2"))
        Console.WriteLine(String.Join("-", hexes))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0A-0F-FF"]);
}

#[test]
fn test_vb_array_convertall_converter_delegate_method_reference() {
    let src = r#"
Imports System

Module Converter
    Public Function DoubleVal(n As Integer) As Integer
        Return n * 2
    End Function
End Module

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3}
        Dim doubled As Integer() = Array.ConvertAll(numbers, AddressOf Converter.DoubleVal)
        Console.WriteLine(String.Join(",", doubled))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,4,6"]);
}
