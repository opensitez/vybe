use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Reflection PropertyInfo, Indexer Parameters & Get/Set Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_reflection_property_info_get_set_value() {
    let src = r#"
Class Account
    Public Property Owner As String = "Alice"
End Class

Module Program
    Sub Main()
        Dim acc As New Account()
        Dim prop = GetType(Account).GetProperty("Owner")
        Console.WriteLine(prop.GetValue(acc))
        prop.SetValue(acc, "Bob")
        Console.WriteLine(acc.Owner)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice", "Bob"]);
}

#[test]
fn test_vb_reflection_property_info_can_read_can_write() {
    let src = r#"
Class ReadOnlyProperty
    Public ReadOnly Property Title As String
    Public Sub New(t As String) : Title = t : End Sub
End Class

Module Program
    Sub Main()
        Dim prop = GetType(ReadOnlyProperty).GetProperty("Title")
        Console.WriteLine(prop.CanRead & "|" & prop.CanWrite)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_reflection_property_info_get_method_set_method() {
    let src = r#"
Imports System.Reflection

Class Container
    Public Property Data As Integer
End Class

Module Program
    Sub Main()
        Dim prop = GetType(Container).GetProperty("Data")
        Dim getMethod = prop.GetGetMethod()
        Dim setMethod = prop.GetSetMethod()
        Console.WriteLine(getMethod.Name & "|" & setMethod.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["get_Data|set_Data"]);
}

#[test]
fn test_vb_reflection_property_info_indexed_property_get_index_parameters() {
    let src = r#"
Imports System.Reflection

Class StringGrid
    Default Public Property Item(row As Integer, col As Integer) As String
        Get
            Return "R" & row & "C" & col
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim prop = GetType(StringGrid).GetProperty("Item")
        Dim indexParams = prop.GetIndexParameters()
        Console.WriteLine(indexParams.Length & ":" & indexParams(0).Name & "," & indexParams(1).Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:row,col"]);
}

#[test]
fn test_vb_reflection_property_info_indexed_property_get_value() {
    let src = r#"
Class SimpleMap
    Default Public Property Item(key As String) As Integer
        Get
            Return key.Length
        End Get
        Set(value As Integer)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim map As New SimpleMap()
        Dim prop = GetType(SimpleMap).GetProperty("Item")
        Dim val = prop.GetValue(map, {"VisualBasic"})
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["11"]);
}

#[test]
fn test_vb_reflection_property_info_indexed_property_set_value() {
    let src = r#"
Imports System.Collections.Generic

Class Cache
    Private store As New Dictionary(Of String, String)()
    Default Public Property Item(key As String) As String
        Get
            Return store(key)
        End Get
        Set(value As String)
            store(key) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim c As New Cache()
        Dim prop = GetType(Cache).GetProperty("Item")
        prop.SetValue(c, "Data100", {"K1"})
        Console.WriteLine(c("K1"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Data100"]);
}

#[test]
fn test_vb_reflection_property_info_shared_static_property() {
    let src = r#"
Class SystemState
    Public Shared Property AppName As String = "VybeApp"
End Class

Module Program
    Sub Main()
        Dim prop = GetType(SystemState).GetProperty("AppName")
        Console.WriteLine(prop.GetValue(Nothing))
        prop.SetValue(Nothing, "NewVybeApp")
        Console.WriteLine(SystemState.AppName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["VybeApp", "NewVybeApp"]);
}

#[test]
fn test_vb_reflection_property_info_private_setter() {
    let src = r#"
Imports System.Reflection

Class Config
    Public Property Mode As String
        Get
            Return "Production"
        End Get
        Private Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim prop = GetType(Config).GetProperty("Mode")
        Dim pubSet = prop.GetSetMethod(False)
        Dim nonPubSet = prop.GetSetMethod(True)
        Console.WriteLine((pubSet Is Nothing) & "|" & (nonPubSet IsNot Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_reflection_property_info_property_type() {
    let src = r#"
Class Sample
    Public Property Count As Integer
    Public Property Tag As String
End Class

Module Program
    Sub Main()
        Dim p1 = GetType(Sample).GetProperty("Count")
        Dim p2 = GetType(Sample).GetProperty("Tag")
        Console.WriteLine(p1.PropertyType.Name & "|" & p2.PropertyType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32|String"]);
}

#[test]
fn test_vb_reflection_property_info_get_custom_attributes() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Property)>
Class RequiredAttribute
    Inherits Attribute
End Class

Class Model
    <Required>
    Public Property Name As String
End Class

Module Program
    Sub Main()
        Dim prop = GetType(Model).GetProperty("Name")
        Dim attrs = prop.GetCustomAttributes(GetType(RequiredAttribute), False)
        Console.WriteLine(attrs.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_reflection_property_info_generic_type_property() {
    let src = r#"
Class Wrapper(Of T)
    Public Property Value As T
    Public Sub New(v As T) : Value = v : End Sub
End Class

Module Program
    Sub Main()
        Dim w As New Wrapper(Of Double)(3.14)
        Dim prop = GetType(Wrapper(Of Double)).GetProperty("Value")
        Console.WriteLine(prop.PropertyType.Name & "=" & prop.GetValue(w))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Double=3.14"]);
}

#[test]
fn test_vb_reflection_property_info_overloaded_indexed_properties() {
    let src = r#"
Class MultiIndexer
    Default Public Property Item(i As Integer) As String
        Get : Return "Int_" & i : End Get
        Set(value As String) : End Set
    End Property
    Default Public Property Item(s As String) As String
        Get : Return "Str_" & s : End Get
        Set(value As String) : End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim mi As New MultiIndexer()
        Dim pInt = GetType(MultiIndexer).GetProperty("Item", {GetType(Integer)})
        Dim pStr = GetType(MultiIndexer).GetProperty("Item", {GetType(String)})

        Console.WriteLine(pInt.GetValue(mi, {5}) & "|" & pStr.GetValue(mi, {"abc"}))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int_5|Str_abc"]);
}

#[test]
fn test_vb_reflection_property_info_virtual_overridden_property() {
    let src = r#"
Class BaseClass
    Public Overridable Property Title As String = "Base"
End Class

Class DerivedClass
    Inherits BaseClass
    Public Overrides Property Title As String = "Derived"
End Class

Module Program
    Sub Main()
        Dim d As New DerivedClass()
        Dim prop = GetType(DerivedClass).GetProperty("Title")
        Console.WriteLine(prop.GetValue(d))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Derived"]);
}

#[test]
fn test_vb_reflection_property_info_enum_property() {
    let src = r#"
Enum Level
    Low
    High
End Enum

Class TaskItem
    Public Property Priority As Level = Level.High
End Class

Module Program
    Sub Main()
        Dim item As New TaskItem()
        Dim prop = GetType(TaskItem).GetProperty("Priority")
        Dim val As Level = CType(prop.GetValue(item), Level)
        Console.WriteLine(val.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["High"]);
}

#[test]
fn test_vb_reflection_property_info_struct_value_type() {
    let src = r#"
Structure Point
    Public Property X As Integer
    Public Property Y As Integer
End Structure

Module Program
    Sub Main()
        Dim pt As Object = New Point With {.X = 10, .Y = 20}
        Dim prop = GetType(Point).GetProperty("X")
        prop.SetValue(pt, 50)
        Dim unboxed As Point = CType(pt, Point)
        Console.WriteLine(unboxed.X)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_reflection_property_info_nullable_property() {
    let src = r#"
Imports System

Class Document
    Public Property Pages As Nullable(Of Integer) = 42
End Class

Module Program
    Sub Main()
        Dim doc As New Document()
        Dim prop = GetType(Document).GetProperty("Pages")
        Dim val = prop.GetValue(doc)
        Console.WriteLine(val.GetType().Name & "=" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32=42"]);
}

#[test]
fn test_vb_reflection_property_info_tuple_property() {
    let src = r#"
Class Entity
    Public Property Pair As (X As Integer, Y As Integer) = (10, 20)
End Class

Module Program
    Sub Main()
        Dim e As New Entity()
        Dim prop = GetType(Entity).GetProperty("Pair")
        Dim tuple As (Integer, Integer) = CType(prop.GetValue(e), (Integer, Integer))
        Console.WriteLine(tuple.Item1 & "," & tuple.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_reflection_property_info_get_properties_binding_flags() {
    let src = r#"
Imports System.Reflection

Class FilterTest
    Public Property P1 As Integer
    Private Property P2 As String
    Public Shared Property P3 As Double
End Class

Module Program
    Sub Main()
        Dim props = GetType(FilterTest).GetProperties(BindingFlags.Instance Or BindingFlags.Public)
        Console.WriteLine(props.Length & ":" & props(0).Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:P1"]);
}

#[test]
fn test_vb_reflection_property_info_declaring_type() {
    let src = r#"
Class Parent
    Public Property Tag As String
End Class

Class Child : Inherits Parent : End Class

Module Program
    Sub Main()
        Dim prop = GetType(Child).GetProperty("Tag")
        Console.WriteLine(prop.DeclaringType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Parent"]);
}

#[test]
fn test_vb_reflection_property_info_set_value_null_reference_target_throws() {
    let src = r#"
Imports System
Imports System.Reflection

Class TargetClass
    Public Property Text As String
End Class

Module Program
    Sub Main()
        Dim prop = GetType(TargetClass).GetProperty("Text")
        Try
            prop.SetValue(Nothing, "Val")
        Catch ex As TargetException
            Console.WriteLine("TargetException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TargetException Caught"]);
}
