//! Reflection: GetType, PropertyInfo, MethodInfo, FieldInfo, Activator.
use super::helpers::run_csharp;

#[test]
fn get_type_returns_runtime_type_of_instance() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(42.GetType().Name);"#),
        &["Int32"]
    );
}

#[test]
fn typeof_on_string_has_correct_full_name() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(typeof(string).FullName);"#),
        &["System.String"]
    );
}

#[test]
fn get_properties_lists_public_properties_of_class() {
    assert_eq!(
        run_csharp(
            r#"class Item { public int Id {get;set;} public string Name {get;set;} }
Console.WriteLine(typeof(Item).GetProperties().Length);"#
        ),
        &["2"]
    );
}

#[test]
fn property_info_get_value_reads_property_dynamically() {
    assert_eq!(
        run_csharp(
            r#"class Item { public int Id {get;set;} }
var item = new Item { Id=7 };
var prop = typeof(Item).GetProperty("Id");
Console.WriteLine(prop.GetValue(item));"#
        ),
        &["7"]
    );
}

#[test]
fn property_info_set_value_writes_property_dynamically() {
    assert_eq!(
        run_csharp(
            r#"class Item { public int Id {get;set;} }
var item = new Item();
var prop = typeof(Item).GetProperty("Id");
prop.SetValue(item, 99);
Console.WriteLine(item.Id);"#
        ),
        &["99"]
    );
}

#[test]
fn method_info_invoke_calls_method_dynamically() {
    assert_eq!(
        run_csharp(
            r#"class Calc { public int Double(int n) => n * 2; }
var obj = new Calc();
var method = typeof(Calc).GetMethod("Double");
Console.WriteLine(method.Invoke(obj, new object[]{5}));"#
        ),
        &["10"]
    );
}

#[test]
fn field_info_get_value_reads_public_field() {
    assert_eq!(
        run_csharp(
            r#"class Data { public int X = 3; }
var obj = new Data();
var field = typeof(Data).GetField("X");
Console.WriteLine(field.GetValue(obj));"#
        ),
        &["3"]
    );
}

#[test]
fn activator_create_instance_constructs_parameterless_type() {
    assert_eq!(
        run_csharp(
            r#"class Widget { public int Value = 42; }
var w = (Widget)System.Activator.CreateInstance(typeof(Widget));
Console.WriteLine(w.Value);"#
        ),
        &["42"]
    );
}

#[test]
fn get_methods_count_includes_public_instance_methods() {
    assert_eq!(
        run_csharp(
            r#"class Calc { public int Add(int a, int b) => a+b; public int Sub(int a, int b) => a-b; }
Console.WriteLine(typeof(Calc).GetMethods(
    System.Reflection.BindingFlags.Public|System.Reflection.BindingFlags.Instance|
    System.Reflection.BindingFlags.DeclaredOnly).Length);"#
        ),
        &["2"]
    );
}
