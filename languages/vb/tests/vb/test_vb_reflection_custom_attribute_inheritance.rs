use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom Attributes Inheritance & Reflection Discovery
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_custom_attribute_basic_class_decoration() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class)>
Class AuthorAttribute
    Inherits Attribute
    Public Name As String
    Public Sub New(n As String) : Name = n : End Sub
End Class

<Author("Alice")>
Class Document : End Class

Module Program
    Sub Main()
        Dim t = GetType(Document)
        Dim attr = CType(t.GetCustomAttributes(GetType(AuthorAttribute), False)(0), AuthorAttribute)
        Console.WriteLine(attr.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice"]);
}

#[test]
fn test_vb_custom_attribute_inherited_flag_true() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class, Inherited:=True)>
Class CategoryAttribute
    Inherits Attribute
    Public Tag As String
    Public Sub New(t As String) : Tag = t : End Sub
End Class

<Category("BaseTag")>
Class BaseClass : End Class

Class DerivedClass : Inherits BaseClass : End Class

Module Program
    Sub Main()
        Dim t = GetType(DerivedClass)
        Dim attrs = t.GetCustomAttributes(GetType(CategoryAttribute), True)
        Console.WriteLine(attrs.Length & ":" & CType(attrs(0), CategoryAttribute).Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:BaseTag"]);
}

#[test]
fn test_vb_custom_attribute_inherited_flag_false() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class, Inherited:=False)>
Class NonInheritedAttribute
    Inherits Attribute
End Class

<NonInherited>
Class BaseClass : End Class

Class DerivedClass : Inherits BaseClass : End Class

Module Program
    Sub Main()
        Dim attrs = GetType(DerivedClass).GetCustomAttributes(GetType(NonInheritedAttribute), True)
        Console.WriteLine(attrs.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_custom_attribute_allow_multiple_true() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class, AllowMultiple:=True)>
Class TagAttribute
    Inherits Attribute
    Public Tag As String
    Public Sub New(t As String) : Tag = t : End Sub
End Class

<Tag("Alpha")>
<Tag("Beta")>
Class Item : End Class

Module Program
    Sub Main()
        Dim attrs = GetType(Item).GetCustomAttributes(GetType(TagAttribute), False)
        Console.WriteLine(attrs.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_custom_attribute_named_positional_arguments() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class)>
Class MetadataAttribute
    Inherits Attribute
    Public ReadOnly ID As Integer
    Public Property Description As String
    Public Property Version As Integer = 1
    Public Sub New(idVal As Integer) : ID = idVal : End Sub
End Class

<Metadata(100, Description:="ServiceClass", Version:=2)>
Class Service : End Class

Module Program
    Sub Main()
        Dim attr = CType(GetType(Service).GetCustomAttributes(GetType(MetadataAttribute), False)(0), MetadataAttribute)
        Console.WriteLine(attr.ID & "|" & attr.Description & "|v" & attr.Version)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100|ServiceClass|v2"]);
}

#[test]
fn test_vb_custom_attribute_applied_to_interface() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Interface)>
Class ContractAttribute
    Inherits Attribute
    Public Name As String
    Public Sub New(n As String) : Name = n : End Sub
End Class

<Contract("IServiceContract")>
Interface IService : End Interface

Module Program
    Sub Main()
        Dim attr = CType(GetType(IService).GetCustomAttributes(GetType(ContractAttribute), False)(0), ContractAttribute)
        Console.WriteLine(attr.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IServiceContract"]);
}

#[test]
fn test_vb_custom_attribute_applied_to_enum() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Enum)>
Class BitFlagsAttribute
    Inherits Attribute
End Class

<BitFlags>
Enum Privileges
    Read = 1
    Write = 2
End Enum

Module Program
    Sub Main()
        Dim isDefined = GetType(Privileges).IsDefined(GetType(BitFlagsAttribute), False)
        Console.WriteLine(isDefined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_custom_attribute_applied_to_struct() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Struct)>
Class ImmutableAttribute
    Inherits Attribute
End Class

<Immutable>
Structure Vector2D
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Sub Main()
        Dim isDefined = GetType(Vector2D).IsDefined(GetType(ImmutableAttribute), False)
        Console.WriteLine(isDefined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_custom_attribute_applied_to_parameter() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Parameter)>
Class NonNullAttribute
    Inherits Attribute
End Class

Class Validator
    Public Sub Process(<NonNull> data As String) : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(Validator).GetMethod("Process")
        Dim p = m.GetParameters()(0)
        Dim isDefined = p.IsDefined(GetType(NonNullAttribute), False)
        Console.WriteLine(isDefined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_custom_attribute_applied_to_return_value() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.ReturnValue)>
Class NonNullReturnAttribute
    Inherits Attribute
End Class

Class Factory
    Public Function Create() As String : Return "Item" : End Function
End Class

Module Program
    Sub Main()
        Dim m = GetType(Factory).GetMethod("Create")
        Dim returnParam = m.ReturnParameter
        Console.WriteLine(returnParam IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_custom_attribute_applied_to_assembly() {
    let src = r#"
Imports System
Imports System.Reflection

<Assembly: AssemblyTitle("MyVybeAssembly")>

Module Program
    Sub Main()
        Dim asm = Assembly.GetExecutingAssembly()
        Dim titleAttr = CType(asm.GetCustomAttributes(GetType(AssemblyTitleAttribute), False)(0), AssemblyTitleAttribute)
        Console.WriteLine(titleAttr.Title)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["MyVybeAssembly"]);
}

#[test]
fn test_vb_custom_attribute_applied_to_module() {
    let src = r#"
Imports System
Imports System.Reflection

<Module: Description("VybeModule")>

Module Program
    Sub Main()
        Dim modInfo = GetType(Program).Module
        Dim isDefined = modInfo.IsDefined(GetType(DescriptionAttribute), False)
        Console.WriteLine(isDefined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_custom_attribute_derived_attribute_class() {
    let src = r#"
Imports System

Class BaseAttribute
    Inherits Attribute
    Public Message As String
    Public Sub New(msg As String) : Message = msg : End Sub
End Class

Class SpecializedAttribute
    Inherits BaseAttribute
    Public Sub New(msg As String) : MyBase.New("Spec: " & msg) : End Sub
End Class

<Specialized("CustomNote")>
Class TargetClass : End Class

Module Program
    Sub Main()
        Dim attr = CType(GetType(TargetClass).GetCustomAttributes(GetType(BaseAttribute), True)(0), BaseAttribute)
        Console.WriteLine(attr.Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Spec: CustomNote"]);
}

#[test]
fn test_vb_custom_attribute_typeof_parameter() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class)>
Class RelatedTypeAttribute
    Inherits Attribute
    Public TargetType As Type
    Public Sub New(t As Type) : TargetType = t : End Sub
End Class

<RelatedType(GetType(String))>
Class StringProcessor : End Class

Module Program
    Sub Main()
        Dim attr = CType(GetType(StringProcessor).GetCustomAttributes(GetType(RelatedTypeAttribute), False)(0), RelatedTypeAttribute)
        Console.WriteLine(attr.TargetType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["String"]);
}

#[test]
fn test_vb_custom_attribute_enum_parameter() {
    let src = r#"
Imports System

Enum LogLevel
    Debug
    Info
    ErrorVal
End Enum

<AttributeUsage(AttributeTargets.Method)>
Class LogAttribute
    Inherits Attribute
    Public Level As LogLevel
    Public Sub New(l As LogLevel) : Level = l : End Sub
End Class

Class Service
    <Log(LogLevel.ErrorVal)>
    Public Sub Process() : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(Service).GetMethod("Process")
        Dim attr = CType(m.GetCustomAttributes(GetType(LogAttribute), False)(0), LogAttribute)
        Console.WriteLine(attr.Level.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ErrorVal"]);
}

#[test]
fn test_vb_custom_attribute_array_parameter() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class)>
Class RolesAttribute
    Inherits Attribute
    Public Roles As String()
    Public Sub New(ParamArray r As String()) : Roles = r : End Sub
End Class

<Roles("Admin", "Manager")>
Class Dashboard : End Class

Module Program
    Sub Main()
        Dim attr = CType(GetType(Dashboard).GetCustomAttributes(GetType(RolesAttribute), False)(0), RolesAttribute)
        Console.WriteLine(String.Join(",", attr.Roles))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Admin,Manager"]);
}

#[test]
fn test_vb_custom_attribute_get_custom_attributes_data() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class)>
Class TagAttribute
    Inherits Attribute
    Public Sub New(tag As String) : End Sub
End Class

<Tag("SampleTag")>
Class AnnotatedClass : End Class

Module Program
    Sub Main()
        Dim customData = GetType(AnnotatedClass).GetCustomAttributesData()
        Console.WriteLine(customData.Count & ":" & customData(0).ConstructorArguments(0).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:SampleTag"]);
}

#[test]
fn test_vb_custom_attribute_method_overrides_inheritance() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Method, Inherited:=True)>
Class TraceAttribute
    Inherits Attribute
End Class

Class BaseService
    <Trace>
    Public Overridable Sub DoWork() : End Sub
End Class

Class DerivedService
    Inherits BaseService
    Public Overrides Sub DoWork() : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(DerivedService).GetMethod("DoWork")
        Dim isDefined = m.IsDefined(GetType(TraceAttribute), True)
        Console.WriteLine(isDefined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_custom_attribute_generic_class_decoration() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class)>
Class GenericTagAttribute
    Inherits Attribute
End Class

<GenericTag>
Class Repository(Of T) : End Class

Module Program
    Sub Main()
        Dim isDefined = GetType(Repository(Of String)).IsDefined(GetType(GenericTagAttribute), False)
        Console.WriteLine(isDefined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_custom_attribute_multiple_attributes_combination() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class)>
Class AttrA : Inherits Attribute : End Class

<AttributeUsage(AttributeTargets.Class)>
Class AttrB : Inherits Attribute : End Class

<AttrA>
<AttrB>
Class MultiAnnotated : End Class

Module Program
    Sub Main()
        Dim allAttrs = GetType(MultiAnnotated).GetCustomAttributes(False)
        Console.WriteLine(allAttrs.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}
