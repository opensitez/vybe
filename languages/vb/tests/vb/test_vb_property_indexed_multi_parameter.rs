use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Multi-Parameter Indexed Properties & Matrix Access
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_indexed_property_2d_matrix_getter_setter() {
    let src = r#"
Class Grid
    Private data(2, 2) As Integer
    Default Public Property Cell(r As Integer, c As Integer) As Integer
        Get
            Return data(r, c)
        End Get
        Set(value As Integer)
            data(r, c) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim g As New Grid()
        g(1, 2) = 42
        Console.WriteLine(g(1, 2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_indexed_property_named_non_default() {
    let src = r#"
Class Table
    Private data(1, 1) As String
    Public Property ItemAt(row As Integer, col As Integer) As String
        Get
            Return data(row, col)
        End Get
        Set(value As String)
            data(row, col) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim t As New Table()
        t.ItemAt(0, 1) = "Header"
        Console.WriteLine(t.ItemAt(0, 1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Header"]);
}

#[test]
fn test_vb_indexed_property_string_and_int_keys() {
    let src = r#"
Imports System.Collections.Generic

Class MultiMap
    Private dict As New Dictionary(Of String, List(Of String))()
    Default Public Property Value(category As String, index As Integer) As String
        Get
            If dict.ContainsKey(category) AndAlso index < dict(category).Count Then
                Return dict(category)(index)
            End If
            Return Nothing
        End Get
        Set(val As String)
            If Not dict.ContainsKey(category) Then
                dict(category) = New List(Of String)()
            End If
            dict(category).Add(val)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim m As New MultiMap()
        m("Fruits", 0) = "Apple"
        m("Fruits", 1) = "Banana"
        Console.WriteLine(m("Fruits", 0) & "|" & m("Fruits", 1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Apple|Banana"]);
}

#[test]
fn test_vb_indexed_property_read_only_3d_cube() {
    let src = r#"
Class Cube
    Public ReadOnly Property Coordinate(x As Integer, y As Integer, z As Integer) As String
        Get
            Return "(" & x & "," & y & "," & z & ")"
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim c As New Cube()
        Console.WriteLine(c.Coordinate(1, 2, 3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["(1,2,3)"]);
}

#[test]
fn test_vb_indexed_property_write_only_log() {
    let src = r#"
Class MemoryLog
    Private logs(2) As String
    Public WriteOnly Property LogEntry(index As Integer) As String
        Set(value As String)
            logs(index) = "LOG: " & value
        End Set
    End Property
    Public Function ReadLog(index As Integer) As String
        Return logs(index)
    End Function
End Class

Module Program
    Sub Main()
        Dim l As New MemoryLog()
        l.LogEntry(0) = "Started"
        Console.WriteLine(l.ReadLog(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["LOG: Started"]);
}

#[test]
fn test_vb_indexed_property_overloading_parameter_types() {
    let src = r#"
Class OverloadedIndexer
    Default Public Property Item(key As String) As String
        Get
            Return "StringKey:" & key
        End Get
        Set(value As String)
        End Set
    End Property

    Default Public Property Item(key As Integer) As String
        Get
            Return "IntKey:" & key
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim idx As New OverloadedIndexer()
        Console.WriteLine(idx("test") & "|" & idx(100))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["StringKey:test|IntKey:100"]);
}

#[test]
fn test_vb_indexed_property_overloading_parameter_count() {
    let src = r#"
Class GridContainer
    Default Public Property Item(x As Integer) As String
        Get
            Return "1D:" & x
        End Get
        Set(value As String)
        End Set
    End Property

    Default Public Property Item(x As Integer, y As Integer) As String
        Get
            Return "2D:" & x & "," & y
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim g As New GridContainer()
        Console.WriteLine(g(5) & "|" & g(5, 10))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1D:5|2D:5,10"]);
}

#[test]
fn test_vb_indexed_property_optional_parameters() {
    let src = r#"
Class OptionalIndexer
    Public Property Element(row As Integer, Optional col As Integer = 0) As String
        Get
            Return "R=" & row & ",C=" & col
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim o As New OptionalIndexer()
        Console.WriteLine(o.Element(5) & "|" & o.Element(5, 3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["R=5,C=0|R=5,C=3"]);
}

#[test]
fn test_vb_indexed_property_paramarray_arguments() {
    let src = r#"
Class VariadicIndexer
    Public Property KeyPath(ParamArray keys As String()) As String
        Get
            Return String.Join("/", keys)
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim v As New VariadicIndexer()
        Console.WriteLine(v.KeyPath("usr", "local", "bin"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["usr/local/bin"]);
}

#[test]
fn test_vb_indexed_property_in_interface() {
    let src = r#"
Interface IMatrix
    Default Property Item(r As Integer, c As Integer) As Double
End Interface

Class DoubleMatrix
    Implements IMatrix
    Private arr(1, 1) As Double
    Default Public Property Item(r As Integer, c As Integer) As Double Implements IMatrix.Item
        Get
            Return arr(r, c)
        End Get
        Set(value As Double)
            arr(r, c) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim m As IMatrix = New DoubleMatrix()
        m(0, 1) = 3.14
        Console.WriteLine(m(0, 1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3.14"]);
}

#[test]
fn test_vb_indexed_property_generic_class() {
    let src = r#"
Class GenericGrid(Of T)
    Private data(1, 1) As T
    Default Public Property Cell(r As Integer, c As Integer) As T
        Get
            Return data(r, c)
        End Get
        Set(value As T)
            data(r, c) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim g As New GenericGrid(Of String)()
        g(0, 0) = "TopLeft"
        Console.WriteLine(g(0, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TopLeft"]);
}

#[test]
fn test_vb_indexed_property_struct_return_mutation() {
    let src = r#"
Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Class PointGrid
    Private points(1) As Point
    Default Public Property Item(idx As Integer) As Point
        Get
            Return points(idx)
        End Get
        Set(value As Point)
            points(idx) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim pg As New PointGrid()
        pg(0) = New Point With {.X = 10, .Y = 20}
        Console.WriteLine(pg(0).X & "," & pg(0).Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_indexed_property_custom_enum_indexer() {
    let src = r#"
Enum Channel
    Red = 0
    Green = 1
    Blue = 2
End Enum

Class ColorPixel
    Private channels(2) As Byte
    Default Public Property Component(ch As Channel) As Byte
        Get
            Return channels(CInt(ch))
        End Get
        Set(value As Byte)
            channels(CInt(ch)) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim p As New ColorPixel()
        p(Channel.Green) = 255
        Console.WriteLine(p(Channel.Green))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255"]);
}

#[test]
fn test_vb_indexed_property_derived_class_shadowing() {
    let src = r#"
Class ParentMap
    Default Public Property Item(key As String) As String
        Get
            Return "Parent:" & key
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Class ChildMap
    Inherits ParentMap
    Default Public Shadows Property Item(key As String) As String
        Get
            Return "Child:" & key
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim c As New ChildMap()
        Dim p As ParentMap = c
        Console.WriteLine(c("test") & "|" & p("test"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Child:test|Parent:test"]);
}

#[test]
fn test_vb_indexed_property_derived_class_override() {
    let src = r#"
Class BaseStore
    Default Public Overridable Property Item(id As Integer) As String
        Get
            Return "BaseStore"
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Class CustomStore
    Inherits BaseStore
    Default Public Overrides Property Item(id As Integer) As String
        Get
            Return "CustomStore_" & id
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim s As BaseStore = New CustomStore()
        Console.WriteLine(s(10))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CustomStore_10"]);
}

#[test]
fn test_vb_indexed_property_tuple_key() {
    let src = r#"
Imports System.Collections.Generic

Class TupleMap
    Private dict As New Dictionary(Of (Integer, Integer), String)()
    Default Public Property Item(r As Integer, c As Integer) As String
        Get
            If dict.ContainsKey((r, c)) Then Return dict((r, c))
            Return Nothing
        End Get
        Set(value As String)
            dict((r, c)) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim tm As New TupleMap()
        tm(3, 4) = "Position34"
        Console.WriteLine(tm(3, 4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Position34"]);
}

#[test]
fn test_vb_indexed_property_side_effects_on_getter() {
    let src = r#"
Class AccessCounter
    Private _count As Integer = 0
    Default Public Property Item(idx As Integer) As String
        Get
            _count += 1
            Return "Access_" & _count
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim ac As New AccessCounter()
        Console.WriteLine(ac(0) & "|" & ac(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Access_1|Access_2"]);
}

#[test]
fn test_vb_indexed_property_date_time_indexer() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Class Schedule
    Private events As New Dictionary(Of DateTime, String)()
    Default Public Property EventName(dt As DateTime) As String
        Get
            If events.ContainsKey(dt) Then Return events(dt)
            Return "Free"
        Get
        End Get
        Set(value As String)
            events(dt) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim s As New Schedule()
        Dim dt = New DateTime(2025, 6, 1)
        s(dt) = "Conference"
        Console.WriteLine(s(dt) & "|" & s(dt.AddDays(1)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Conference|Free"]);
}

#[test]
fn test_vb_indexed_property_compound_assignment_operator() {
    let src = r#"
Class NumericGrid
    Private arr(2) As Integer
    Default Public Property Value(i As Integer) As Integer
        Get
            Return arr(i)
        End Get
        Set(val As Integer)
            arr(i) = val
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim g As New NumericGrid()
        g(0) = 10
        g(0) += 5
        Console.WriteLine(g(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15"]);
}

#[test]
fn test_vb_indexed_property_reference_type_null_checks() {
    let src = r#"
Class NullableStore
    Private items(1) As String
    Default Public Property Item(i As Integer) As String
        Get
            Return items(i)
        End Get
        Set(val As String)
            items(i) = val
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim ns As New NullableStore()
        Console.WriteLine(ns(0) Is Nothing)
        ns(0) = "Set"
        Console.WriteLine(ns(0) Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}
