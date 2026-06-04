use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(
    typeof_full_name_reports_primitive_runtime_name,
    r#"Console.WriteLine(typeof(int).FullName);"#,
    ["System.Int32"]
);
csharp_case!(
    typeof_name_reports_string_type_name,
    r#"Console.WriteLine(typeof(string).Name);"#,
    ["String"]
);
csharp_case!(
    gettype_on_instance_reports_runtime_type_name,
    r#"object value = "hello"; Console.WriteLine(value.GetType().Name);"#,
    ["String"]
);
csharp_case!(
    activator_create_instance_builds_default_constructed_object,
    r#"using System; class Box { public int Value = 4; } var box = (Box)Activator.CreateInstance(typeof(Box)); Console.WriteLine(box.Value);"#,
    ["4"]
);
csharp_case!(
    constructor_info_can_invoke_parameterized_constructor,
    r#"using System; class Box { public string Name; public Box(string name) { Name = name; } } var ctor = typeof(Box).GetConstructor(new[] { typeof(string) }); var box = (Box)ctor.Invoke(new object[] { "crate" }); Console.WriteLine(box.Name);"#,
    ["crate"]
);
csharp_case!(
    method_info_invokes_instance_method_and_returns_value,
    r#"using System; class Box { public string Read() { return "value"; } } var method = typeof(Box).GetMethod("Read"); Console.WriteLine(method.Invoke(new Box(), null));"#,
    ["value"]
);
csharp_case!(
    property_info_reads_property_value_from_instance,
    r#"using System; class Box { public string Name { get; set; } = "pkg"; } var prop = typeof(Box).GetProperty("Name"); Console.WriteLine(prop.GetValue(new Box()));"#,
    ["pkg"]
);
csharp_case!(
    field_info_reads_public_field_value_from_instance,
    r#"using System; class Box { public int Count = 12; } var field = typeof(Box).GetField("Count"); Console.WriteLine(field.GetValue(new Box()));"#,
    ["12"]
);
csharp_case!(
    property_info_sets_property_value_on_instance,
    r#"using System; class Box { public string Name { get; set; } } var box = new Box(); var prop = typeof(Box).GetProperty("Name"); prop.SetValue(box, "updated"); Console.WriteLine(box.Name);"#,
    ["updated"]
);
csharp_case!(
    field_info_sets_public_field_value_on_instance,
    r#"using System; class Box { public int Count; } var box = new Box(); var field = typeof(Box).GetField("Count"); field.SetValue(box, 9); Console.WriteLine(box.Count);"#,
    ["9"]
);
csharp_case!(
    is_assignable_from_reports_true_for_derived_type,
    r#"class Base { } class Child : Base { } Console.WriteLine(typeof(Base).IsAssignableFrom(typeof(Child)));"#,
    ["True"]
);
csharp_case!(
    is_value_type_reports_true_for_struct_type,
    r#"Console.WriteLine(typeof(System.DateTime).IsValueType);"#,
    ["True"]
);
csharp_case!(
    base_type_reports_parent_class_name,
    r#"class Base { } class Child : Base { } Console.WriteLine(typeof(Child).BaseType.Name);"#,
    ["Base"]
);
csharp_case!(
    get_nested_type_finds_declared_inner_class,
    r#"class Outer { public class Inner { } } Console.WriteLine(typeof(Outer).GetNestedType("Inner") != null);"#,
    ["True"]
);
csharp_case!(
    get_generic_arguments_reports_type_argument_name,
    r#"class Box<T> { } Console.WriteLine(typeof(Box<int>).GetGenericArguments()[0].Name);"#,
    ["Int32"]
);
csharp_case!(
    get_generic_type_definition_returns_open_generic_name,
    r#"class Box<T> { } Console.WriteLine(typeof(Box<int>).GetGenericTypeDefinition().Name.Contains("Box"));"#,
    ["True"]
);
csharp_case!(
    property_info_can_report_can_write_for_settable_property,
    r#"using System; class Box { public string Name { get; set; } } Console.WriteLine(typeof(Box).GetProperty("Name").CanWrite);"#,
    ["True"]
);
csharp_case!(
    method_info_reports_static_for_static_method,
    r#"using System; class Box { public static string Read() { return "ok"; } } Console.WriteLine(typeof(Box).GetMethod("Read").IsStatic);"#,
    ["True"]
);
csharp_case!(
    enum_type_reports_isenum_true,
    r#"enum State { Ready } Console.WriteLine(typeof(State).IsEnum);"#,
    ["True"]
);
csharp_case!(
    interface_list_contains_declared_contract_name,
    r#"using System.Linq; interface IRun { } class Worker : IRun { } var names = typeof(Worker).GetInterfaces().Select(i => i.Name); foreach (var name in names) Console.WriteLine(name);"#,
    ["IRun"]
);
