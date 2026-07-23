use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Reflection FieldInfo GetValue, SetValue & Flags
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_reflection_field_info_public_instance_get_set() {
    let src = r#"
Class DataBox
    Public Tag As String = "Initial"
End Class

Module Program
    Sub Main()
        Dim box As New DataBox()
        Dim field = GetType(DataBox).GetField("Tag")
        Console.WriteLine(field.GetValue(box))
        field.SetValue(box, "UpdatedTag")
        Console.WriteLine(box.Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Initial", "UpdatedTag"]);
}

#[test]
fn test_vb_reflection_field_info_private_instance_binding_flags() {
    let src = r#"
Imports System.Reflection

Class Account
    Private _balance As Double = 500.0
    Public Function GetBalance() As Double : Return _balance : End Function
End Class

Module Program
    Sub Main()
        Dim acc As New Account()
        Dim field = GetType(Account).GetField("_balance", BindingFlags.Instance Or BindingFlags.NonPublic)
        Console.WriteLine(field.GetValue(acc))
        field.SetValue(acc, 1000.0)
        Console.WriteLine(acc.GetBalance())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["500", "1000"]);
}

#[test]
fn test_vb_reflection_field_info_public_shared_static() {
    let src = r#"
Class GlobalConfig
    Public Shared Version As String = "1.0.0"
End Class

Module Program
    Sub Main()
        Dim field = GetType(GlobalConfig).GetField("Version")
        Console.WriteLine(field.GetValue(Nothing))
        field.SetValue(Nothing, "2.0.0")
        Console.WriteLine(GlobalConfig.Version)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.0.0", "2.0.0"]);
}

#[test]
fn test_vb_reflection_field_info_readonly_field_check() {
    let src = r#"
Class ReadOnlyContainer
    Public ReadOnly ID As Integer = 42
    Public Sub New(idVal As Integer) : ID = idVal : End Sub
End Class

Module Program
    Sub Main()
        Dim field = GetType(ReadOnlyContainer).GetField("ID")
        Console.WriteLine(field.IsInitOnly)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_reflection_field_info_literal_const_field() {
    let src = r#"
Class Constants
    Public Const MaxUsers As Integer = 100
End Class

Module Program
    Sub Main()
        Dim field = GetType(Constants).GetField("MaxUsers")
        Console.WriteLine(field.IsLiteral & "|" & field.GetRawConstantValue())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|100"]);
}

#[test]
fn test_vb_reflection_field_info_value_type_struct_boxing() {
    let src = r#"
Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Sub Main()
        Dim pt As Object = New Point With {.X = 10, .Y = 20}
        Dim fieldX = GetType(Point).GetField("X")
        fieldX.SetValue(pt, 99)
        Dim unboxed As Point = CType(pt, Point)
        Console.WriteLine(unboxed.X)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}

#[test]
fn test_vb_reflection_field_info_get_fields_filter() {
    let src = r#"
Imports System.Reflection

Class MultiField
    Public F1 As Integer
    Public F2 As String
    Private F3 As Double
End Class

Module Program
    Sub Main()
        Dim pubFields = GetType(MultiField).GetFields(BindingFlags.Instance Or BindingFlags.Public)
        Dim allFields = GetType(MultiField).GetFields(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        Console.WriteLine(pubFields.Length & "|" & allFields.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|3"]);
}

#[test]
fn test_vb_reflection_field_info_field_type_property() {
    let src = r#"
Class Entity
    Public Title As String
    Public Age As Integer
End Class

Module Program
    Sub Main()
        Dim f1 = GetType(Entity).GetField("Title")
        Dim f2 = GetType(Entity).GetField("Age")
        Console.WriteLine(f1.FieldType.Name & "|" & f2.FieldType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["String|Int32"]);
}

#[test]
fn test_vb_reflection_field_info_is_public_is_private() {
    let src = r#"
Imports System.Reflection

Class Security
    Public OpenData As String
    Private SecretData As String
End Class

Module Program
    Sub Main()
        Dim fPub = GetType(Security).GetField("OpenData")
        Dim fPriv = GetType(Security).GetField("SecretData", BindingFlags.Instance Or BindingFlags.NonPublic)
        Console.WriteLine(fPub.IsPublic & "|" & fPriv.IsPrivate)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_reflection_field_info_generic_class_fields() {
    let src = r#"
Class GenericHolder(Of T)
    Public Element As T
End Class

Module Program
    Sub Main()
        Dim hInt As New GenericHolder(Of Integer) With {.Element = 42}
        Dim field = GetType(GenericHolder(Of Integer)).GetField("Element")
        Console.WriteLine(field.FieldType.Name & "=" & field.GetValue(hInt))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32=42"]);
}

#[test]
fn test_vb_reflection_field_info_custom_attribute() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Field)>
Class RangeAttribute
    Inherits Attribute
    Public Min As Integer
    Public Max As Integer
    Public Sub New(min As Integer, max As Integer) : Me.Min = min : Me.Max = max : End Sub
End Class

Class Form
    <Range(1, 100)>
    Public Percentage As Integer
End Class

Module Program
    Sub Main()
        Dim field = GetType(Form).GetField("Percentage")
        Dim attr = CType(field.GetCustomAttributes(GetType(RangeAttribute), False)(0), RangeAttribute)
        Console.WriteLine(attr.Min & " To " & attr.Max)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1 To 100"]);
}

#[test]
fn test_vb_reflection_field_info_enum_field() {
    let src = r#"
Enum Level
    Low = 1
    High = 2
End Enum

Module Program
    Sub Main()
        Dim field = GetType(Level).GetField("High")
        Console.WriteLine(field.IsLiteral & "|" & field.GetRawConstantValue())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|2"]);
}

#[test]
fn test_vb_reflection_field_info_inherited_fields() {
    let src = r#"
Class BaseClass
    Public BaseField As String = "Base"
End Class

Class DerivedClass
    Inherits BaseClass
    Public DerivedField As String = "Derived"
End Class

Module Program
    Sub Main()
        Dim d As New DerivedClass()
        Dim fBase = GetType(DerivedClass).GetField("BaseField")
        Console.WriteLine(fBase.GetValue(d))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Base"]);
}

#[test]
fn test_vb_reflection_field_info_declaring_type() {
    let src = r#"
Class Parent
    Public Value As Integer
End Class

Class Child : Inherits Parent : End Class

Module Program
    Sub Main()
        Dim field = GetType(Child).GetField("Value")
        Console.WriteLine(field.DeclaringType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Parent"]);
}

#[test]
fn test_vb_reflection_field_info_set_value_type_mismatch_throws() {
    let src = r#"
Imports System

Class Holder
    Public Num As Integer
End Class

Module Program
    Sub Main()
        Dim h As New Holder()
        Dim field = GetType(Holder).GetField("Num")
        Try
            field.SetValue(h, "NotAnInteger")
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException on Type Mismatch Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentException on Type Mismatch Caught"]
    );
}

#[test]
fn test_vb_reflection_field_info_nullable_type() {
    let src = r#"
Imports System

Class Container
    Public Score As Nullable(Of Double)
End Class

Module Program
    Sub Main()
        Dim c As New Container()
        Dim field = GetType(Container).GetField("Score")
        field.SetValue(c, 98.5)
        Console.WriteLine(c.Score.HasValue & ":" & c.Score.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:98.5"]);
}

#[test]
fn test_vb_reflection_field_info_tuple_field() {
    let src = r#"
Class TupleHolder
    Public Pair As (String, Integer) = ("A", 1)
End Class

Module Program
    Sub Main()
        Dim th As New TupleHolder()
        Dim field = GetType(TupleHolder).GetField("Pair")
        Dim val As (String, Integer) = CType(field.GetValue(th), (String, Integer))
        Console.WriteLine(val.Item1 & "=" & val.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A=1"]);
}

#[test]
fn test_vb_reflection_field_info_array_field() {
    let src = r#"
Class ArrayHolder
    Public Items As String() = {"X", "Y"}
End Class

Module Program
    Sub Main()
        Dim ah As New ArrayHolder()
        Dim field = GetType(ArrayHolder).GetField("Items")
        Dim arr As String() = CType(field.GetValue(ah), String())
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X,Y"]);
}

#[test]
fn test_vb_reflection_field_info_is_static_property() {
    let src = r#"
Imports System.Reflection

Class MemberMix
    Public Inst As Integer
    Public Shared Stat As Integer
End Class

Module Program
    Sub Main()
        Dim fInst = GetType(MemberMix).GetField("Inst")
        Dim fStat = GetType(MemberMix).GetField("Stat")
        Console.WriteLine(fInst.IsStatic & "|" & fStat.IsStatic)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|True"]);
}

#[test]
fn test_vb_reflection_field_info_set_value_direct_reference_mutation() {
    let src = r#"
Class Node
    Public NextNode As Node
    Public Name As String
    Public Sub New(n As String) : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim n1 As New Node("N1")
        Dim n2 As New Node("N2")
        Dim field = GetType(Node).GetField("NextNode")
        field.SetValue(n1, n2)
        Console.WriteLine(n1.NextNode.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["N2"]);
}
