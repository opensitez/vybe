use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Reflection ConstructorInfo & Activator.CreateInstance
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_reflection_activator_create_instance_parameterless() {
    let src = r#"
Imports System

Class Config
    Public Status As String = "Ready"
End Class

Module Program
    Sub Main()
        Dim t = GetType(Config)
        Dim instance As Config = CType(Activator.CreateInstance(t), Config)
        Console.WriteLine(instance.Status)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Ready"]);
}

#[test]
fn test_vb_reflection_activator_create_instance_with_args() {
    let src = r#"
Imports System

Class User
    Public Property Name As String
    Public Property Age As Integer
    Public Sub New(n As String, a As Integer)
        Name = n : Age = a
    End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(User)
        Dim u As User = CType(Activator.CreateInstance(t, "Alice", 30), User)
        Console.WriteLine(u.Name & " is " & u.Age)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice is 30"]);
}

#[test]
fn test_vb_reflection_constructor_info_invoke() {
    let src = r#"
Imports System
Imports System.Reflection

Class Product
    Public SKU As String
    Public Sub New(s As String) : SKU = s : End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(Product)
        Dim ctor = t.GetConstructor({GetType(String)})
        Dim p As Product = CType(ctor.Invoke({"SKU-100"}), Product)
        Console.WriteLine(p.SKU)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SKU-100"]);
}

#[test]
fn test_vb_reflection_constructor_info_get_parameters() {
    let src = r#"
Imports System
Imports System.Reflection

Class Order
    Public Sub New(id As Integer, title As String) : End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(Order)
        Dim ctor = t.GetConstructors()(0)
        Dim params = ctor.GetParameters()
        Console.WriteLine(params(0).Name & ":" & params(0).ParameterType.Name & "|" & params(1).Name & ":" & params(1).ParameterType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["id:Int32|title:String"]);
}

#[test]
fn test_vb_reflection_private_constructor_invoke_binding_flags() {
    let src = r#"
Imports System
Imports System.Reflection

Class Singleton
    Private Sub New()
        Console.WriteLine("Private Constructor Invoked")
    End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(Singleton)
        Dim ctor = t.GetConstructor(BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, Type.EmptyTypes, Nothing)
        Dim instance = ctor.Invoke(Nothing)
        Console.WriteLine(instance IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Private Constructor Invoked", "True"]);
}

#[test]
fn test_vb_reflection_activator_create_instance_generic_type() {
    let src = r#"
Imports System

Class Container(Of T)
    Public Item As T
    Public Sub New(i As T) : Item = i : End Sub
End Class

Module Program
    Sub Main()
        Dim genericType = GetType(Container(Of )).MakeGenericType(GetType(String))
        Dim instance = Activator.CreateInstance(genericType, "GenericData")
        Dim prop = genericType.GetField("Item")
        Console.WriteLine(prop.GetValue(instance))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["GenericData"]);
}

#[test]
fn test_vb_reflection_get_constructors_count() {
    let src = r#"
Class MultiCtor
    Public Sub New() : End Sub
    Public Sub New(a As Integer) : End Sub
    Public Sub New(a As Integer, b As String) : End Sub
End Class

Module Program
    Sub Main()
        Dim ctors = GetType(MultiCtor).GetConstructors()
        Console.WriteLine(ctors.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_reflection_activator_create_instance_value_type_struct() {
    let src = r#"
Imports System

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Structure

Module Program
    Sub Main()
        Dim pt As Point = CType(Activator.CreateInstance(GetType(Point), 5, 10), Point)
        Console.WriteLine(pt.X & "," & pt.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5,10"]);
}

#[test]
fn test_vb_reflection_activator_create_instance_enum_type() {
    let src = r#"
Imports System

Enum Status
    Active = 1
End Enum

Module Program
    Sub Main()
        Dim s = CType(Activator.CreateInstance(GetType(Status)), Status)
        Console.WriteLine(s.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_reflection_constructor_is_public_is_private() {
    let src = r#"
Imports System
Imports System.Reflection

Class Sample
    Public Sub New() : End Sub
    Private Sub New(x As Integer) : End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(Sample)
        Dim pubCtor = t.GetConstructor(Type.EmptyTypes)
        Dim privCtor = t.GetConstructor(BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, {GetType(Integer)}, Nothing)

        Console.WriteLine(pubCtor.IsPublic & "|" & privCtor.IsPrivate)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_reflection_constructor_params_byref() {
    let src = r#"
Imports System
Imports System.Reflection

Class ByRefCtor
    Public Sub New(ByRef val As Integer)
        val = 999
    End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(ByRefCtor)
        Dim ctor = t.GetConstructors()(0)
        Dim args As Object() = {10}
        ctor.Invoke(args)
        Console.WriteLine(args(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["999"]);
}

#[test]
fn test_vb_reflection_activator_create_instance_type_name_string() {
    let src = r#"
Imports System

Class Target
    Public Function Echo() As String : Return "EchoTarget" : End Function
End Class

Module Program
    Sub Main()
        Dim handle = Activator.CreateInstance(Nothing, "Target")
        Dim obj As Target = CType(handle.Unwrap(), Target)
        Console.WriteLine(obj.Echo())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["EchoTarget"]);
}

#[test]
fn test_vb_reflection_constructor_attributes() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Constructor)>
Class TagAttribute
    Inherits Attribute
    Public Note As String
    Public Sub New(n As String) : Note = n : End Sub
End Class

Class Annotated
    <Tag("PrimaryCtor")>
    Public Sub New() : End Sub
End Class

Module Program
    Sub Main()
        Dim ctor = GetType(Annotated).GetConstructor(Type.EmptyTypes)
        Dim attr = CType(ctor.GetCustomAttributes(GetType(TagAttribute), False)(0), TagAttribute)
        Console.WriteLine(attr.Note)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["PrimaryCtor"]);
}

#[test]
fn test_vb_reflection_constructor_throws_target_invocation_exception() {
    let src = r#"
Imports System
Imports System.Reflection

Class ThrowingCtor
    Public Sub New()
        Throw New InvalidOperationException("Ctor Failed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(ThrowingCtor)
        Try
            Activator.CreateInstance(t)
        Catch ex As TargetInvocationException
            Console.WriteLine(ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Ctor Failed"]);
}

#[test]
fn test_vb_reflection_constructor_info_contains_generic_parameters() {
    let src = r#"
Class GenericCtor(Of T)
    Public Sub New(item As T) : End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(GenericCtor(Of String))
        Dim ctor = t.GetConstructors()(0)
        Console.WriteLine(ctor.GetParameters()(0).ParameterType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["String"]);
}

#[test]
fn test_vb_reflection_static_constructor_is_not_in_get_constructors() {
    let src = r#"
Class WithStaticCtor
    Shared Sub New() : End Sub
    Public Sub New() : End Sub
End Class

Module Program
    Sub Main()
        Dim ctors = GetType(WithStaticCtor).GetConstructors()
        Console.WriteLine(ctors.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_reflection_type_initializer_static_constructor_info() {
    let src = r#"
Imports System.Reflection

Class StaticCtorClass
    Shared Sub New() : End Sub
End Class

Module Program
    Sub Main()
        Dim ctor = GetType(StaticCtorClass).TypeInitializer
        Console.WriteLine(ctor.IsStatic)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_reflection_activator_create_instance_array_type() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim arr = CType(Activator.CreateInstance(GetType(String()), 3), String())
        Console.WriteLine(arr.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_reflection_constructor_info_calling_convention() {
    let src = r#"
Class Sample : End Class

Module Program
    Sub Main()
        Dim ctor = GetType(Sample).GetConstructor(Type.EmptyTypes)
        Console.WriteLine(ctor.CallingConvention.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["HasThis"]);
}

#[test]
fn test_vb_reflection_activator_create_instance_non_public_flag() {
    let src = r#"
Imports System

Class InternalClass
    Private Sub New() : End Sub
End Class

Module Program
    Sub Main()
        Dim instance = Activator.CreateInstance(GetType(InternalClass), True)
        Console.WriteLine(instance IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
