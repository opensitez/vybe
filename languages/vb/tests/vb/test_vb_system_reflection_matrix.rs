use super::helpers::run_vb;

#[test]
fn reflection_get_type_shape_metadata() {
    let out = run_vb(
        r#"
Imports System
Imports System.Reflection

Module M
    Sub Main()
        Dim t As Type = GetType(System.Text.StringBuilder)
        Console.WriteLine(t.IsClass)
        Console.WriteLine(t.Name)
        Console.WriteLine(t.Namespace)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "StringBuilder", "System.Text"]);
}

#[test]
fn reflection_public_property_lookup() {
    let out = run_vb(
        r#"
Imports System
Imports System.Reflection

Module M
    Sub Main()
        Dim t As Type = GetType(TestType)
        Dim prop As PropertyInfo = t.GetProperty("Value")
        Dim obj As New TestType()
        prop.SetValue(obj, 9)
        Console.WriteLine(prop.GetValue(obj))
    End Sub

    Class TestType
        Public Property Value As Integer
    End Class
End Module
"#,
    );

    assert_eq!(out, vec!["9"]);
}

#[test]
fn reflection_get_fields_and_methods() {
    let out = run_vb(
        r#"
Imports System
Imports System.Reflection

Module M
    Sub Main()
        Dim t As Type = GetType(Composite)
        Console.WriteLine(t.GetFields().Length >= 1)
        Console.WriteLine(t.GetMethods().Length >= 1)
    End Sub

    Class Composite
        Public X As Integer
        Public Y As Integer
        Public Sub Inc()
            X += 1
        End Sub
        Public Sub Dec()
            X -= 1
        End Sub
    End Class
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn reflection_invoke_method_dynamically() {
    let out = run_vb(
        r#"
Imports System
Imports System.Reflection

Module M
    Sub Main()
        Dim t As Type = GetType(Operation)
        Dim obj As New Operation()
        Dim method As MethodInfo = t.GetMethod("Add")
        Dim result As Object = method.Invoke(obj, New Object(){3, 4})
        Console.WriteLine(result)
    End Sub

    Class Operation
        Public Function Add(a As Integer, b As Integer) As Integer
            Return a + b
        End Function
    End Class
End Module
"#,
    );

    assert_eq!(out, vec!["7"]);
}

#[test]
fn reflection_binding_flags_controls_visibility() {
    let out = run_vb(
        r#"
Imports System
Imports System.Reflection

Module M
    Sub Main()
        Dim t As Type = GetType(Container)
        Dim fields = t.GetFields(BindingFlags.Instance Or BindingFlags.NonPublic)
        Dim p As PropertyInfo = t.GetProperty("Visible", BindingFlags.Instance Or BindingFlags.Public)
        Console.WriteLine(fields.Length)
        Console.WriteLine(p.Name)
    End Sub

    Class Container
        Private Secret As Integer = 7
        Public Property Visible As Integer
            Get
                Return Secret
            End Get
            Set
                Secret = Value
            End Set
        End Property
    End Class
End Module
"#,
    );

    assert_eq!(out, vec!["1", "Visible"]);
}

#[test]
fn reflection_interfaces_discovery() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim impl As New Target()
        Dim types() As Type = impl.GetType().GetInterfaces()
        Console.WriteLine(types.Length)
        Console.WriteLine(types(0).Name)
    End Sub

    Interface ITraceable
        Sub Mark()
    End Interface

    Class Target
        Implements ITraceable
        Public Sub Mark() Implements ITraceable.Mark
        End Sub
    End Class
End Module
"#,
    );

    assert_eq!(out, vec!["1", "ITraceable"]);
}

#[test]
fn reflection_activation_with_constructor_and_fields() {
    let out = run_vb(
        r#"
Imports System
Imports System.Reflection

Module M
    Sub Main()
        Dim t As Type = GetType(Buildable)
        Dim obj As Object = Activator.CreateInstance(t)
        Dim ctor As ConstructorInfo = t.GetConstructor(Type.EmptyTypes)
        Console.WriteLine(ctor IsNot Nothing)
        Dim value As Integer = CType(t.GetField("Counter").GetValue(obj), Integer)
        Console.WriteLine(value)
        Console.WriteLine(obj.GetType().Name)
    End Sub

    Class Buildable
        Public Counter As Integer = 11
    End Class
End Module
"#,
    );

    assert_eq!(out, vec!["True", "11", "Buildable"]);
}
