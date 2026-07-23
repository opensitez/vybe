use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ValueTuple Deconstruct Overloads & Custom Deconstruction
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_tuple_deconstruct_2_elements() {
    let src = r#"
Module Program
    Sub Main()
        Dim tuple As (String, Integer) = ("Alice", 30)
        Dim name As String = Nothing
        Dim age As Integer = 0
        tuple.Deconstruct(name, age)
        Console.WriteLine(name & " is " & age)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice is 30"]);
}

#[test]
fn test_vb_tuple_deconstruct_3_elements() {
    let src = r#"
Module Program
    Sub Main()
        Dim tuple As (Integer, Double, String) = (100, 99.9, "PASSED")
        Dim id As Integer = 0
        Dim score As Double = 0.0
        Dim status As String = Nothing
        tuple.Deconstruct(id, score, status)
        Console.WriteLine(id & "|" & score & "|" & status)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100|99.9|PASSED"]);
}

#[test]
fn test_vb_tuple_deconstruct_custom_class() {
    let src = r#"
Class Person
    Public Property Name As String
    Public Property Age As Integer
    Public Sub New(n As String, a As Integer) : Name = n : Age = a : End Sub
    Public Sub Deconstruct(ByRef n As String, ByRef a As Integer)
        n = Name : a = Age
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Person("Bob", 40)
        Dim n As String = Nothing
        Dim a As Integer = 0
        p.Deconstruct(n, a)
        Console.WriteLine(n & " is " & a)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bob is 40"]);
}

#[test]
fn test_vb_tuple_deconstruct_extension_method() {
    let src = r#"
Imports System.Runtime.CompilerServices

Class Point2D
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Class

Module PointExtensions
    <Extension()>
    Public Sub Deconstruct(p As Point2D, ByRef x As Integer, ByRef y As Integer)
        x = p.X : y = p.Y
    End Sub
End Module

Module Program
    Sub Main()
        Dim pt As New Point2D(15, 25)
        Dim x As Integer = 0
        Dim y As Integer = 0
        pt.Deconstruct(x, y)
        Console.WriteLine(x & "," & y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15,25"]);
}

#[test]
fn test_vb_tuple_deconstruct_overloaded_parameter_counts() {
    let src = r#"
Class DateTimeInfo
    Public Sub Deconstruct(ByRef year As Integer, ByRef month As Integer)
        year = 2025 : month = 12
    End Sub
    Public Sub Deconstruct(ByRef year As Integer, ByRef month As Integer, ByRef day As Integer)
        year = 2025 : month = 12 : day = 31
    End Sub
End Class

Module Program
    Sub Main()
        Dim info As New DateTimeInfo()
        Dim y As Integer = 0, m As Integer = 0, d As Integer = 0
        info.Deconstruct(y, m)
        Console.WriteLine(y & "-" & m)
        info.Deconstruct(y, m, d)
        Console.WriteLine(y & "-" & m & "-" & d)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-12", "2025-12-31"]);
}

#[test]
fn test_vb_tuple_literal_syntax_named() {
    let src = r#"
Module Program
    Sub Main()
        Dim pt = (X:=10, Y:=20)
        Console.WriteLine(pt.X & ":" & pt.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10:20"]);
}

#[test]
fn test_vb_tuple_equality_operators() {
    let src = r#"
Module Program
    Sub Main()
        Dim t1 = (1, "A")
        Dim t2 = (1, "A")
        Dim t3 = (2, "B")
        Console.WriteLine((t1 = t2) & "|" & (t1 <> t3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_tuple_comparison_operators() {
    let src = r#"
Module Program
    Sub Main()
        Dim t1 = (1, 10)
        Dim t2 = (1, 20)
        Dim t3 = (2, 5)
        Console.WriteLine((t1 < t2) & "|" & (t2 < t3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_tuple_return_value_from_function() {
    let src = r#"
Module Program
    Private Function GetCoordinates() As (X As Integer, Y As Integer)
        Return (100, 200)
    End Function

    Sub Main()
        Dim coord = GetCoordinates()
        Console.WriteLine(coord.X & "," & coord.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100,200"]);
}

#[test]
fn test_vb_tuple_byref_argument_passing() {
    let src = r#"
Module Program
    Private Sub ModifyTuple(ByRef t As (Integer, String))
        t.Item1 = 99
        t.Item2 = "Updated"
    End Sub

    Sub Main()
        Dim t = (10, "Original")
        ModifyTuple(t)
        Console.WriteLine(t.Item1 & ":" & t.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99:Updated"]);
}

#[test]
fn test_vb_tuple_array_declaration_and_iteration() {
    let src = r#"
Module Program
    Sub Main()
        Dim pairs As (Key As String, Val As Integer)() = {("A", 1), ("B", 2)}
        For Each p In pairs
            Console.WriteLine(p.Key & "=" & p.Val)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A=1", "B=2"]);
}

#[test]
fn test_vb_tuple_nested_in_list() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of (Code As Integer, Message As String)) From {
            (200, "OK"),
            (404, "Not Found")
        }
        Console.WriteLine(list(0).Code & " " & list(0).Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["200 OK"]);
}

#[test]
fn test_vb_tuple_deconstruct_into_existing_variables() {
    let src = r#"
Module Program
    Sub Main()
        Dim t = (100, "User")
        Dim id As Integer = 0
        Dim role As String = Nothing
        t.Deconstruct(id, role)
        Console.WriteLine("ID: " & id & " | Role: " & role)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ID: 100 | Role: User"]);
}

#[test]
fn test_vb_tuple_deconstruct_with_null_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim t As (String, Integer) = (Nothing, 0)
        Dim s As String = "Initial"
        Dim i As Integer = -1
        t.Deconstruct(s, i)
        Console.WriteLine((s Is Nothing) & "|" & i)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|0"]);
}

#[test]
fn test_vb_tuple_4_elements_deconstruct() {
    let src = r#"
Module Program
    Sub Main()
        Dim t = (1, 2, 3, 4)
        Dim a As Integer = 0, b As Integer = 0, c As Integer = 0, d As Integer = 0
        t.Deconstruct(a, b, c, d)
        Console.WriteLine(a + b + c + d)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_tuple_generic_deconstruct() {
    let src = r#"
Class Container(Of T1, T2)
    Public V1 As T1
    Public V2 As T2
    Public Sub New(v1 As T1, v2 As T2) : Me.V1 = v1 : Me.V2 = v2 : End Sub
    Public Sub Deconstruct(ByRef out1 As T1, ByRef out2 As T2)
        out1 = V1 : out2 = V2
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Container(Of String, Double)("PI", 3.14)
        Dim k As String = Nothing
        Dim v As Double = 0.0
        c.Deconstruct(k, v)
        Console.WriteLine(k & "=" & v)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["PI=3.14"]);
}

#[test]
fn test_vb_tuple_struct_custom_deconstruct() {
    let src = r#"
Structure Dimensions
    Public Width As Integer
    Public Height As Integer
    Public Sub New(w As Integer, h As Integer) : Width = w : Height = h : End Sub
    Public Sub Deconstruct(ByRef w As Integer, ByRef h As Integer)
        w = Width : h = Height
    End Sub
End Structure

Module Program
    Sub Main()
        Dim d As New Dimensions(1920, 1080)
        Dim w As Integer = 0, h As Integer = 0
        d.Deconstruct(w, h)
        Console.WriteLine(w & "x" & h)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1920x1080"]);
}

#[test]
fn test_vb_tuple_create_factory_method() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim vt = ValueTuple.Create("Alpha", 100)
        Console.WriteLine(vt.Item1 & ":" & vt.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alpha:100"]);
}

#[test]
fn test_vb_tuple_to_string_representation() {
    let src = r#"
Module Program
    Sub Main()
        Dim t = (1, "Two", True)
        Console.WriteLine(t.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["(1, Two, True)"]);
}

#[test]
fn test_vb_tuple_hash_code_equality() {
    let src = r#"
Module Program
    Sub Main()
        Dim t1 = ("K", 42)
        Dim t2 = ("K", 42)
        Console.WriteLine(t1.GetHashCode() = t2.GetHashCode())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
