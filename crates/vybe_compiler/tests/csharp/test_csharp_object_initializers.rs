//! Object initializers, collection initializers, and nested initializer syntax.
use super::helpers::run_csharp;

#[test]
fn object_initializer_sets_multiple_properties() {
    assert_eq!(
        run_csharp(r#"class Person{public string Name;public int Age;}
var p=new Person{Name="Alice",Age=30};
Console.WriteLine(p.Name); Console.WriteLine(p.Age);"#),
        &["Alice", "30"]
    );
}

#[test]
fn collection_initializer_populates_list_inline() {
    assert_eq!(
        run_csharp(r#"var list=new System.Collections.Generic.List<int>{10,20,30};
Console.WriteLine(list[1]);"#),
        &["20"]
    );
}

#[test]
fn dictionary_initializer_uses_key_value_syntax() {
    assert_eq!(
        run_csharp(r#"var d=new System.Collections.Generic.Dictionary<string,int>{{"a",1},{"b",2}};
Console.WriteLine(d["b"]);"#),
        &["2"]
    );
}

#[test]
fn nested_object_initializer_sets_inner_object() {
    assert_eq!(
        run_csharp(r#"class Address{public string City;}
class Person{public string Name;public Address Home;}
var p=new Person{Name="Bob",Home=new Address{City="Paris"}};
Console.WriteLine(p.Home.City);"#),
        &["Paris"]
    );
}

#[test]
fn array_initializer_infers_element_type() {
    assert_eq!(
        run_csharp(r#"var arr=new[]{1,2,3};
Console.WriteLine(arr.GetType().IsArray); Console.WriteLine(arr.Length);"#),
        &["True", "3"]
    );
}

#[test]
fn anonymous_type_initializer_infers_property_names() {
    assert_eq!(
        run_csharp(r#"string name="Alice"; int age=30;
var anon=new{name,age};
Console.WriteLine(anon.name); Console.WriteLine(anon.age);"#),
        &["Alice", "30"]
    );
}

#[test]
fn list_of_objects_with_object_initializers() {
    assert_eq!(
        run_csharp(r#"class Item{public int Id;}
var items=new System.Collections.Generic.List<Item>{new Item{Id=1},new Item{Id=2}};
Console.WriteLine(items[1].Id);"#),
        &["2"]
    );
}

#[test]
fn with_expression_on_nominal_record_is_object_initializer() {
    assert_eq!(
        run_csharp(r#"record Config{public int Port{get;init;}=80;}
var cfg=new Config() with{Port=443};
Console.WriteLine(cfg.Port);"#),
        &["443"]
    );
}
