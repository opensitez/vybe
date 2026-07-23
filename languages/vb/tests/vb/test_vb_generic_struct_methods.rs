use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generic Structs & Generic Methods Surface Area
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_struct_construction_and_fields() {
    let src = r#"
Structure Point(Of T)
    Public X As T
    Public Y As T
    Public Sub New(x As T, y As T)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim p As New Point(Of Integer)(10, 20)
        Console.WriteLine(p.X & "," & p.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_generic_struct_method_with_type_inference() {
    let src = r#"
Structure Pair(Of T1, T2)
    Public Item1 As T1
    Public Item2 As T2
    Public Sub New(i1 As T1, i2 As T2)
        Item1 = i1 : Item2 = i2
    End Sub
    Public Function Swap() As Pair(Of T2, T1)
        Return New Pair(Of T2, T1)(Item2, Item1)
    End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Pair(Of String, Integer)("Age", 30)
        Dim s = p.Swap()
        Console.WriteLine(s.Item1 & ":" & s.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30:Age"]);
}

#[test]
fn test_vb_generic_struct_readonly_property() {
    let src = r#"
Structure OptionVal(Of T)
    Private _val As T
    Private _hasVal As Boolean
    Public ReadOnly Property Value As T
        Get
            Return _val
        End Get
    End Property
    Public ReadOnly Property HasValue As Boolean
        Get
            Return _hasVal
        End Get
    End Property
    Public Sub New(val As T)
        _val = val
        _hasVal = True
    End Sub
End Structure

Module Program
    Sub Main()
        Dim opt As New OptionVal(Of String)("Hello")
        Console.WriteLine(opt.HasValue & "|" & opt.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|Hello"]);
}

#[test]
fn test_vb_generic_struct_interface_implementation() {
    let src = r#"
Interface IContainer(Of T)
    Function GetElement() As T
End Interface

Structure Box(Of T)
    Implements IContainer(Of T)
    Public Element As T
    Public Sub New(e As T)
        Element = e
    End Sub
    Public Function GetElement() As T Implements IContainer(Of T).GetElement
        Return Element
    End Function
End Structure

Module Program
    Sub Main()
        Dim b As IContainer(Of Double) = New Box(Of Double)(3.14)
        Console.WriteLine(b.GetElement())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3.14"]);
}

#[test]
fn test_vb_generic_struct_value_type_pass_by_val_copy() {
    let src = r#"
Structure MutableBox(Of T)
    Public Item As T
    Public Sub New(i As T)
        Item = i
    End Sub
End Structure

Module Program
    Private Sub ModifyBox(b As MutableBox(Of Integer))
        b.Item = 99
    End Sub

    Sub Main()
        Dim b As New MutableBox(Of Integer)(10)
        ModifyBox(b)
        Console.WriteLine(b.Item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_generic_struct_byref_parameter_mutation() {
    let src = r#"
Structure MutableBox(Of T)
    Public Item As T
    Public Sub New(i As T)
        Item = i
    End Sub
End Structure

Module Program
    Private Sub ModifyBox(ByRef b As MutableBox(Of Integer))
        b.Item = 99
    End Sub

    Sub Main()
        Dim b As New MutableBox(Of Integer)(10)
        ModifyBox(b)
        Console.WriteLine(b.Item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}

#[test]
fn test_vb_generic_struct_static_shared_field() {
    let src = r#"
Structure StaticStruct(Of T)
    Public Shared DefaultVal As T
    Public Item As T
    Public Sub New(i As T)
        Item = i
    End Sub
End Structure

Module Program
    Sub Main()
        StaticStruct(Of Integer).DefaultVal = -1
        StaticStruct(Of String).DefaultVal = "N/A"

        Console.WriteLine(StaticStruct(Of Integer).DefaultVal & "|" & StaticStruct(Of String).DefaultVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1|N/A"]);
}

#[test]
fn test_vb_generic_struct_override_tostring() {
    let src = r#"
Structure Vector2D(Of T)
    Public X As T
    Public Y As T
    Public Sub New(x As T, y As T)
        Me.X = x : Me.Y = y
    End Sub
    Public Overrides Function ToString() As String
        Return "[" & X.ToString() & ", " & Y.ToString() & "]"
    End Function
End Structure

Module Program
    Sub Main()
        Dim v As New Vector2D(Of Single)(1.5F, 2.5F)
        Console.WriteLine(v.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[1.5, 2.5]"]);
}

#[test]
fn test_vb_generic_struct_override_equals_hashcode() {
    let src = r#"
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
        Console.WriteLine(t1.Equals(t2) & "|" & t1.Equals(t3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_generic_struct_constraint_icomparable() {
    let src = r#"
Imports System

Structure Range(Of T As IComparable(Of T))
    Public Min As T
    Public Max As T
    Public Sub New(min As T, max As T)
        Me.Min = min : Me.Max = max
    End Sub
    Public Function Contains(val As T) As Boolean
        Return val.CompareTo(Min) >= 0 AndAlso val.CompareTo(Max) <= 0
    End Function
End Structure

Module Program
    Sub Main()
        Dim r As New Range(Of Integer)(10, 20)
        Console.WriteLine(r.Contains(15) & "|" & r.Contains(25))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_generic_struct_array_allocation() {
    let src = r#"
Structure Cell(Of T)
    Public Value As T
    Public Sub New(v As T) : Value = v : End Sub
End Structure

Module Program
    Sub Main()
        Dim cells(2) As Cell(Of Integer)
        cells(0) = New Cell(Of Integer)(100)
        cells(1) = New Cell(Of Integer)(200)
        Console.WriteLine(cells(0).Value & "+" & cells(1).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["300"]);
}

#[test]
fn test_vb_generic_struct_nested_in_generic_class() {
    let src = r#"
Class Container(Of T)
    Public Structure Entry
        Public Key As String
        Public Value As T
        Public Sub New(k As String, v As T)
            Key = k : Value = v
        End Sub
    End Structure
End Class

Module Program
    Sub Main()
        Dim e As New Container(Of Integer).Entry("Age", 25)
        Console.WriteLine(e.Key & "=" & e.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Age=25"]);
}

#[test]
fn test_vb_generic_struct_default_value_keyword() {
    let src = r#"
Structure Holder(Of T)
    Public Item As T
    Public Function GetDefault() As T
        Return Nothing ' VB "Nothing" for generics evaluates to default(T)
    End Function
End Structure

Module Program
    Sub Main()
        Dim hInt As New Holder(Of Integer)()
        Dim hStr As New Holder(Of String)()
        Console.WriteLine(hInt.GetDefault() & "|" & (hStr.GetDefault() Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|True"]);
}

#[test]
fn test_vb_generic_struct_tuple_field() {
    let src = r#"
Structure TupleHolder(Of T)
    Public Data As (Key As String, Value As T)
    Public Sub New(k As String, v As T)
        Data = (k, v)
    End Sub
End Structure

Module Program
    Sub Main()
        Dim th As New TupleHolder(Of Double)("PI", 3.14159)
        Console.WriteLine(th.Data.Key & "=" & th.Data.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["PI=3.14159"]);
}

#[test]
fn test_vb_generic_struct_enum_type_arg() {
    let src = r#"
Enum State
    Off = 0
    OnVal = 1
End Enum

Structure StateWrapper(Of T As Structure)
    Public CurrentState As T
    Public Sub New(s As T)
        CurrentState = s
    End Sub
End Structure

Module Program
    Sub Main()
        Dim w As New StateWrapper(Of State)(State.OnVal)
        Console.WriteLine(w.CurrentState.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OnVal"]);
}

#[test]
fn test_vb_generic_struct_constructor_overloads() {
    let src = r#"
Structure FlexBox(Of T)
    Public Val As T
    Public Name As String
    Public Sub New(v As T)
        Val = v : Name = "Unnamed"
    End Sub
    Public Sub New(v As T, n As String)
        Val = v : Name = n
    End Sub
End Structure

Module Program
    Sub Main()
        Dim b1 As New FlexBox(Of Integer)(10)
        Dim b2 As New FlexBox(Of Integer)(20, "Custom")
        Console.WriteLine(b1.Name & "|" & b2.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Unnamed|Custom"]);
}

#[test]
fn test_vb_generic_struct_generic_method_inside_generic_struct() {
    let src = r#"
Structure ConverterStruct(Of TInput)
    Public InputData As TInput
    Public Sub New(input As TInput)
        InputData = input
    End Sub
    Public Function ConvertTo(Of TOutput)(converter As System.Func(Of TInput, TOutput)) As TOutput
        Return converter(InputData)
    End Function
End Structure

Module Program
    Sub Main()
        Dim cs As New ConverterStruct(Of String)("123")
        Dim res As Integer = cs.ConvertTo(Of Integer)(Function(s) Integer.Parse(s))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123"]);
}

#[test]
fn test_vb_generic_struct_unboxing_cast() {
    let src = r#"
Structure SimpleBox(Of T)
    Public Value As T
    Public Sub New(v As T)
        Value = v
    End Sub
End Structure

Module Program
    Sub Main()
        Dim box As Object = New SimpleBox(Of Integer)(42)
        Dim unboxed = CType(box, SimpleBox(Of Integer))
        Console.WriteLine(unboxed.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_generic_struct_nullable_field() {
    let src = r#"
Imports System

Structure NullableHolder(Of T As Structure)
    Public Value As Nullable(Of T)
    Public Sub New(v As T)
        Value = v
    End Sub
End Structure

Module Program
    Sub Main()
        Dim nh As New NullableHolder(Of Integer)(55)
        Console.WriteLine(nh.Value.HasValue & ":" & nh.Value.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:55"]);
}

#[test]
fn test_vb_generic_struct_ref_returns_simulation() {
    let src = r#"
Structure CounterStruct(Of T)
    Public Value As T
    Public Sub Increment(byRefTarget As ByRefHolder(Of T), incrementFunc As System.Func(Of T, T))
        byRefTarget.Value = incrementFunc(byRefTarget.Value)
    End Sub
End Structure

Class ByRefHolder(Of T)
    Public Property Value As T
End Class

Module Program
    Sub Main()
        Dim cs As New CounterStruct(Of Integer)()
        Dim holder As New ByRefHolder(Of Integer)() With {.Value = 10}
        cs.Increment(holder, Function(n) n + 5)
        Console.WriteLine(holder.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15"]);
}
