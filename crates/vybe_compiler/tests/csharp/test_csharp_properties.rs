//! Auto-properties, expression-bodied, init-only, computed properties.
use super::helpers::run_csharp;

#[test]
fn auto_property_get_set_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"class Person { public string Name { get; set; } }
var p = new Person(); p.Name = "Alice";
Console.WriteLine(p.Name);"#
        ),
        &["Alice"]
    );
}

#[test]
fn auto_property_with_default_value_initializer() {
    assert_eq!(
        run_csharp(
            r#"class Config { public int Timeout { get; set; } = 30; }
Console.WriteLine(new Config().Timeout);"#
        ),
        &["30"]
    );
}

#[test]
fn init_only_property_set_in_object_initializer() {
    assert_eq!(
        run_csharp(
            r#"class Point { public int X { get; init; } public int Y { get; init; } }
var p = new Point { X=1, Y=2 };
Console.WriteLine(p.X); Console.WriteLine(p.Y);"#
        ),
        &["1", "2"]
    );
}

#[test]
fn computed_read_only_property_derived_from_field() {
    assert_eq!(
        run_csharp(
            r#"class Circle { public double Radius; public double Area => System.Math.PI * Radius * Radius; }
Console.WriteLine(System.Math.Round(new Circle{Radius=0}.Area));"#
        ),
        &["0"]
    );
}

#[test]
fn private_setter_prevents_external_mutation() {
    assert_eq!(
        run_csharp(
            r#"class Counter {
    public int Count { get; private set; }
    public void Increment() => Count++;
}
var c = new Counter(); c.Increment(); c.Increment();
Console.WriteLine(c.Count);"#
        ),
        &["2"]
    );
}

#[test]
fn expression_bodied_property_getter() {
    assert_eq!(
        run_csharp(
            r#"class Rect { public int W,H; public int Area => W * H; }
Console.WriteLine(new Rect{W=3,H=4}.Area);"#
        ),
        &["12"]
    );
}

#[test]
fn backing_field_used_in_custom_setter_logic() {
    assert_eq!(
        run_csharp(
            r#"class Temperature {
    private double _celsius;
    public double Celsius {
        get => _celsius;
        set => _celsius = value < -273.15 ? -273.15 : value;
    }
}
var t = new Temperature(); t.Celsius = -300;
Console.WriteLine(t.Celsius);"#
        ),
        &["-273.15"]
    );
}

#[test]
fn static_property_shares_value_across_instances() {
    assert_eq!(
        run_csharp(
            r#"class Registry { public static int Count { get; set; } }
Registry.Count = 7;
Console.WriteLine(Registry.Count);"#
        ),
        &["7"]
    );
}
