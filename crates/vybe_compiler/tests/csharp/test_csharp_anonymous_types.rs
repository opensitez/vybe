//! Anonymous type creation, LINQ projection, property access, equality.
use super::helpers::run_csharp;

#[test]
fn anonymous_type_created_with_new_projection_syntax() {
    assert_eq!(
        run_csharp(r#"var a=new{Name="Alice",Age=30};
Console.WriteLine(a.Name); Console.WriteLine(a.Age);"#),
        &["Alice", "30"]
    );
}

#[test]
fn anonymous_type_from_linq_select_projection() {
    assert_eq!(
        run_csharp(r#"var data=new[]{(Id:1,Name:"a"),(Id:2,Name:"b")};
var result=data.Select(d=>new{d.Id,Upper=d.Name.ToUpper()}).ToList();
Console.WriteLine(result[1].Upper);"#),
        &["B"]
    );
}

#[test]
fn two_anonymous_types_with_same_shape_are_equal() {
    assert_eq!(
        run_csharp(r#"var a=new{X=1,Y=2}; var b=new{X=1,Y=2};
Console.WriteLine(a.Equals(b));"#),
        &["True"]
    );
}

#[test]
fn anonymous_type_property_names_inferred_from_variable() {
    assert_eq!(
        run_csharp(r#"int id=7; string name="Bob";
var obj=new{id,name};
Console.WriteLine(obj.id); Console.WriteLine(obj.name);"#),
        &["7", "Bob"]
    );
}

#[test]
fn anonymous_type_to_string_shows_property_values() {
    assert_eq!(
        run_csharp(r#"var a=new{X=3,Y=4};
Console.WriteLine(a.ToString().Contains("X = 3"));"#),
        &["True"]
    );
}
