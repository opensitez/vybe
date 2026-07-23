use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Reflection, Custom Attributes & Dynamic Inspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_custom_attribute_retrieval_on_class() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class)>
Class AuthorAttribute
    Inherits Attribute
    Public ReadOnly Name As String
    Public Sub New(authorName As String)
        Name = authorName
    End Sub
End Class

<Author("Alice")>
Class Document
End Class

Module Program
    Sub Main()
        Dim t = GetType(Document)
        Dim attr = CType(Attribute.GetCustomAttribute(t, GetType(AuthorAttribute)), AuthorAttribute)
        Console.WriteLine(attr.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice"]);
}

#[test]
fn test_vb_custom_attribute_named_parameters() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Property)>
Class ColumnAttribute
    Inherits Attribute
    Public Property Name As String
    Public Property IsKey As Boolean
End Class

Class UserRecord
    <Column(Name:="user_id", IsKey:=True)>
    Public Property UserId As Integer
End Class

Module Program
    Sub Main()
        Dim p = GetType(UserRecord).GetProperty("UserId")
        Dim attr = CType(Attribute.GetCustomAttribute(p, GetType(ColumnAttribute)), ColumnAttribute)
        Console.WriteLine(attr.Name & "|" & attr.IsKey)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["user_id|True"]);
}

#[test]
fn test_vb_reflection_get_properties_and_values() {
    let src = r#"
Imports System
Imports System.Reflection

Class Configuration
    Public Property Host As String = "localhost"
    Public Property Port As Integer = 8080
End Class

Module Program
    Sub Main()
        Dim cfg As New Configuration()
        Dim props = cfg.GetType().GetProperties()
        For Each p In props
            Console.WriteLine(p.Name & "=" & p.GetValue(cfg).ToString())
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Host=localhost", "Port=8080"]);
}

#[test]
fn test_vb_reflection_invoke_method_dynamically() {
    let src = r#"
Imports System

Class Calculator
    Public Function Multiply(a As Integer, b As Integer) As Integer
        Return a * b
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As New Calculator()
        Dim method = calc.GetType().GetMethod("Multiply")
        Dim result = method.Invoke(calc, New Object() {6, 7})
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_reflection_create_instance_activator() {
    let src = r#"
Imports System

Class DynamicPlugin
    Public ReadOnly Name As String = "PluginA"
End Class

Module Program
    Sub Main()
        Dim t = GetType(DynamicPlugin)
        Dim instance = CType(Activator.CreateInstance(t), DynamicPlugin)
        Console.WriteLine(instance.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["PluginA"]);
}

#[test]
fn test_vb_reflection_get_fields_private_binding_flags() {
    let src = r#"
Imports System
Imports System.Reflection

Class SecretHolder
    Private secretCode As String = "Pass123"
End Class

Module Program
    Sub Main()
        Dim sh As New SecretHolder()
        Dim field = sh.GetType().GetField("secretCode", BindingFlags.NonPublic Or BindingFlags.Instance)
        Dim val = CStr(field.GetValue(sh))
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Pass123"]);
}

#[test]
fn test_vb_reflection_generic_type_definition_inspection() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim t = GetType(List(Of String))
        Console.WriteLine(t.IsGenericType & "|" & t.GenericTypeArguments(0).Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|String"]);
}

#[test]
fn test_vb_attribute_multiple_allow_multiple() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class, AllowMultiple:=True)>
Class TagAttribute
    Inherits Attribute
    Public Tag As String
    Public Sub New(t As String)
        Tag = t
    End Sub
End Class

<Tag("V1")>
<Tag("Beta")>
Class Feature
End Class

Module Program
    Sub Main()
        Dim attrs = Attribute.GetCustomAttributes(GetType(Feature), GetType(TagAttribute))
        Dim tags As New System.Collections.Generic.List(Of String)()
        For Each a In attrs
            tags.Add(CType(a, TagAttribute).Tag)
        Next
        Console.WriteLine(String.Join(",", tags))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["V1,Beta"]);
}

#[test]
fn test_vb_reflection_get_events_info() {
    let src = r#"
Imports System
Imports System.Reflection

Class Button
    Public Event Click As EventHandler
End Class

Module Program
    Sub Main()
        Dim events = GetType(Button).GetEvents()
        Console.WriteLine(events(0).Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Click"]);
}

#[test]
fn test_vb_reflection_get_constructors_with_parameters() {
    let src = r#"
Imports System
Imports System.Reflection

Class Person
    Public Sub New(name As String, age As Integer)
    End Sub
End Class

Module Program
    Sub Main()
        Dim ctor = GetType(Person).GetConstructors()(0)
        Dim params = ctor.GetParameters()
        Console.WriteLine(params(0).Name & ":" & params(0).ParameterType.Name & "|" & params(1).Name & ":" & params(1).ParameterType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["name:String|age:Int32"]);
}

#[test]
fn test_vb_reflection_set_value_on_property() {
    let src = r#"
Imports System

Class Account
    Public Property Balance As Decimal
End Class

Module Program
    Sub Main()
        Dim acc As New Account()
        Dim prop = acc.GetType().GetProperty("Balance")
        prop.SetValue(acc, 500.0D)
        Console.WriteLine(acc.Balance)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["500"]);
}

#[test]
fn test_vb_reflection_is_subclass_of_and_assignable_from() {
    let src = r#"
Imports System

Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Module Program
    Sub Main()
        Dim tAnimal = GetType(Animal)
        Dim tDog = GetType(Dog)
        Console.WriteLine(tDog.IsSubclassOf(tAnimal) & "|" & tAnimal.IsAssignableFrom(tDog))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_reflection_make_generic_type_dynamically() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim openType = GetType(List(Of ))
        Dim closedType = openType.MakeGenericType(GetType(Integer))
        Dim listInstance = Activator.CreateInstance(closedType)
        Console.WriteLine(listInstance.GetType().GenericTypeArguments(0).Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32"]);
}

#[test]
fn test_vb_reflection_make_generic_method_dynamically() {
    let src = r#"
Imports System
Imports System.Reflection

Class Utility
    Public Shared Function Wrap(Of T)(val As T) As String
        Return "Wrapped:" & val.ToString()
    End Function
End Class

Module Program
    Sub Main()
        Dim method = GetType(Utility).GetMethod("Wrap")
        Dim genericMethod = method.MakeGenericMethod(GetType(Integer))
        Dim result = genericMethod.Invoke(Nothing, New Object() {99})
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Wrapped:99"]);
}

#[test]
fn test_vb_reflection_assembly_get_executing_assembly() {
    let src = r#"
Imports System.Reflection

Module Program
    Sub Main()
        Dim asm = Assembly.GetExecutingAssembly()
        Console.WriteLine(asm IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_attribute_inherited_search() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class, Inherited:=True)>
Class BaseMetadataAttribute
    Inherits Attribute
    Public Description As String
    Public Sub New(d As String)
        Description = d
    End Sub
End Class

<BaseMetadata("Base Description")>
Class BaseClass
End Class

Class SubClass
    Inherits BaseClass
End Class

Module Program
    Sub Main()
        ' Inherited search enabled (inherit:=True)
        Dim attr = CType(Attribute.GetCustomAttribute(GetType(SubClass), GetType(BaseMetadataAttribute), inherit:=True), BaseMetadataAttribute)
        Console.WriteLine(attr.Description)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Base Description"]);
}

#[test]
fn test_vb_reflection_enum_underlying_type_and_values() {
    let src = r#"
Imports System

Enum Status
    Pending = 10
    Completed = 20
End Enum

Module Program
    Sub Main()
        Dim t = GetType(Status)
        Dim names = [Enum].GetNames(t)
        Dim values = [Enum].GetValues(t)
        Console.WriteLine(String.Join(",", names) & "|Count=" & values.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Pending,Completed|Count=2"]);
}

#[test]
fn test_vb_reflection_get_interfaces_implemented() {
    let src = r#"
Imports System

Interface IA
End Interface

Interface IB
End Interface

Class Service
    Implements IA, IB
End Class

Module Program
    Sub Main()
        Dim interfaces = GetType(Service).GetInterfaces()
        Console.WriteLine(interfaces.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_reflection_delegate_create_delegate() {
    let src = r#"
Imports System
Imports System.Reflection

Class Target
    Public Function MultiplyByTwo(n As Integer) As Integer
        Return n * 2
    End Function
End Class

Module Program
    Sub Main()
        Dim obj As New Target()
        Dim method = obj.GetType().GetMethod("MultiplyByTwo")
        Dim del = CType([Delegate].CreateDelegate(GetType(Func(Of Integer, Integer)), obj, method), Func(Of Integer, Integer))
        Console.WriteLine(del(25))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_reflection_parameter_default_value_inspection() {
    let src = r#"
Imports System

Class Options
    Public Sub Process(Optional timeout As Integer = 30)
    End Sub
End Class

Module Program
    Sub Main()
        Dim param = GetType(Options).GetMethod("Process").GetParameters()(0)
        Console.WriteLine(param.HasDefaultValue & "|" & param.DefaultValue)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|30"]);
}
