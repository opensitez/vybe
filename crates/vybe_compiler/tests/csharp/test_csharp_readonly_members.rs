//! `readonly` fields, `readonly` structs, and `init`-only properties.
use super::helpers::run_csharp;

#[test]
fn readonly_field_set_in_constructor_cannot_change_after() {
    assert_eq!(
        run_csharp(
            r#"class Immutable{public readonly int Value; public Immutable(int v){Value=v;}}
var obj=new Immutable(42);
Console.WriteLine(obj.Value);"#
        ),
        &["42"]
    );
}

#[test]
fn readonly_static_field_initialized_at_class_load() {
    assert_eq!(
        run_csharp(
            r#"class Config{public static readonly string Env="prod";}
Console.WriteLine(Config.Env);"#
        ),
        &["prod"]
    );
}

#[test]
fn readonly_struct_fields_all_readonly_by_definition() {
    assert_eq!(
        run_csharp(
            r#"readonly struct Point{public readonly int X,Y; public Point(int x,int y){X=x;Y=y;}}
var p=new Point(1,2);
Console.WriteLine(p.X+p.Y);"#
        ),
        &["3"]
    );
}

#[test]
fn init_property_settable_only_in_object_initializer() {
    assert_eq!(
        run_csharp(
            r#"class Config{public int Port{get;init;}=80;}
var c=new Config{Port=443};
Console.WriteLine(c.Port);"#
        ),
        &["443"]
    );
}

#[test]
fn record_auto_properties_are_init_by_default() {
    assert_eq!(
        run_csharp(
            r#"record User(string Name,int Age);
var u=new User("Ada",20);
Console.WriteLine(u.Name); Console.WriteLine(u.Age);"#
        ),
        &["Ada", "20"]
    );
}

#[test]
fn const_field_accessible_without_instance_on_type() {
    assert_eq!(
        run_csharp(
            r#"class Physics{public const double C=299792458.0;}
Console.WriteLine(Physics.C>0);"#
        ),
        &["True"]
    );
}

#[test]
fn const_local_not_changeable_but_usable_in_expression() {
    assert_eq!(
        run_csharp(
            r#"const int MAX=100;
Console.WriteLine(MAX*2);"#
        ),
        &["200"]
    );
}
